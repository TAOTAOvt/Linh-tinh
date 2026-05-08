

# nav2_no_gui.py
import os
import sys
import signal
import subprocess
import math
import time
import rclpy
from rclpy.node import Node
from rclpy.executors import MultiThreadedExecutor

import threading
from geometry_msgs.msg import Twist
from nav_msgs.msg import Odometry
from nav_msgs.msg import Path as Pa
from geometry_msgs.msg import PoseStamped

from RS485 import rs_485
from AZDKD import AZDKD as motor
from DIFFROBOT import ROBOT as robot

# ====================== CẤU HÌNH ROBOT ======================
ROBOT_WHEEL_RADIUS = 0.1
ROBOT_WHEEL_SEPARATION = 0.45
ROBOT_MAX_LINEAR_M_S = 40
ROBOT_MIN_LINEAR_M_S = -40
ROBOT_MAX_ANGULAR_R_S = 0.5
ROBOT_MIN_ANGULAR_R_S = -0.5

robot = robot(ROBOT_WHEEL_SEPARATION, ROBOT_WHEEL_RADIUS,
              ROBOT_MAX_LINEAR_M_S, ROBOT_MIN_LINEAR_M_S,
              ROBOT_MAX_ANGULAR_R_S, ROBOT_MIN_ANGULAR_R_S)

az = motor("ContinusOperationWithSpeed")
serial_ = rs_485("/dev/rs485", 1, 8, "E", 115200, 0.5)
client = serial_.connect_()

WHEEL_DIAMETER = 0.20
RPM_TO_MPS = math.pi * WHEEL_DIAMETER / 60.0

# ====================== KILL ROS PROCESSES ======================
def kill_ros_processes_by_name():
    names = [
        "bno055", "ekf_node", "nav2_", "bt_navigator", "planner_server",
        "controller_server", "recoveries_server", "waypoint_follower",
        "static_transform_publisher", "robot_state_publisher",
        "sick_scan_xd", "realsense2_camera"
    ]
    for name in names:
        subprocess.run(["pkill", "-9", "-f", name],
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

# ====================== EXTERNAL NODES ======================
external_ros_processes = []

def launch_external_nodes():
    global external_ros_processes
    procs = []

    procs.append(subprocess.Popen([
        "ros2", "run", "bno055", "bno055",
        "--ros-args", "--params-file","/home/ipx/ros2_foxy_ws/src/bno055/bno055/params/bno055_params.yaml"
    ], preexec_fn=os.setsid))

    procs.append(subprocess.Popen([
        "ros2", "run", "robot_localization", "ekf_node",
        "--ros-args", "--params-file", "/home/ipx/Downloads/Project_tracking/ekf.yaml"
    ], preexec_fn=os.setsid))

    procs.append(subprocess.Popen([
        "ros2", "run", "tf2_ros", "static_transform_publisher",
        "-0.15", "0", "0.0", "0", "0", "0", "base_link", "bno055"
    ], preexec_fn=os.setsid))

    #procs.append(subprocess.Popen([
    #    "ros2", "launch", "rplidar_ros", "rplidar.launch.py"
    #], preexec_fn=os.setsid))

    procs.append(subprocess.Popen([
        "ros2", "launch", "sick_scan_xd", "sick_tim_5xx.launch.py",
        "min_ang:=-0.610865238", "max_ang:=0.610865238",
    ], preexec_fn=os.setsid))

    procs.append(subprocess.Popen([
        "ros2", "launch", "realsense2_camera", "rs_launch.py", "align_depth:=true"
    ], preexec_fn=os.setsid))

    procs.append(subprocess.Popen([
        "ros2", "run", "tf2_ros", "static_transform_publisher",
        "0.22", "0.0", "0.0", "0", "0", "0", "base_link", "uwb_anchor"
    ], preexec_fn=os.setsid))

    procs.append(subprocess.Popen([
        "ros2", "run", "tf2_ros", "static_transform_publisher",
        "0.52", "0.0", "-0.15", "0", "0", "0", "base_link", "camera_link"
    ], preexec_fn=os.setsid))

    #procs.append(subprocess.Popen([
    #    "ros2", "run", "tf2_ros", "static_transform_publisher",
    #    "0", "0", "0.24", "0", "0", "0", "base_link", "laser"
    #], preexec_fn=os.setsid))

    #procs.append(subprocess.Popen([
    #    "ros2", "launch", "nav2_bringup", "navigation_launch.py",
    #    "use_sim_time:=false",
    #    "params_file:=/home/ipx/Downloads/Project_tracking/nav2_params.yaml"
    #], preexec_fn=os.setsid))

    external_ros_processes = procs
    return procs

# ====================== FEEDBACK & CONTROL ======================
def feedback():
    registerSpeedHZ = 0xCE
    vl_feedback = client.read_holding_registers(address=registerSpeedHZ, count=2, slave=0x03)
    encoderL = (vl_feedback.registers[0] << 16) | vl_feedback.registers[1]
    if encoderL >= 0x80000000:
        encoderL -= 0x100000000

    vr_feedback = client.read_holding_registers(address=registerSpeedHZ, count=2, slave=0x02)
    encoderR = (vr_feedback.registers[0] << 16) | vr_feedback.registers[1]
    if encoderR >= 0x80000000:
        encoderR -= 0x100000000

    return encoderL, encoderR

def rpm_to_mps(rpm: float) -> float:
    return rpm * RPM_TO_MPS

def control_robot(linear_velocity, angular_velocity):
    vl, vr = robot.kinematic_(linear_velocity, angular_velocity)
    vl_cmd = az.directComand_(vl)
    vr_cmd = az.directComand_(vr)

    def send_left():
        try:
            client.write_registers(vl_cmd[0], vl_cmd[1], slave=0x03)
        except Exception as e:
            print("[Lỗi motor trái]", e)

    def send_right():
        try:
            client.write_registers(vr_cmd[0], vr_cmd[1], slave=0x02)
        except Exception as e:
            print("[Lỗi motor phải]", e)

    threading.Thread(target=send_left, daemon=True).start()
    threading.Thread(target=send_right, daemon=True).start()

# ====================== ROS NODES ======================
class CmdVelListener(Node):
    def __init__(self):
        super().__init__('cmd_vel_listener')
        self.subscription = self.create_subscription(Twist, '/cmd_vel_dk', self.callback, 10)

    def callback(self, msg: Twist):
        control_robot(msg.linear.x, msg.angular.z)

class OdometryPublisher(Node):
    def __init__(self):
        super().__init__('odom_publisher')
        self.publisher_ = self.create_publisher(Odometry, '/odom_wheel', 10)
        self.timer = self.create_timer(0.1, self.timer_callback)
        self.x = 0.0
        self.y = 0.0
        self.theta = 0.0  # rad
        self.last_encoderL = None
        self.last_encoderR = None
        self.last_time = self.get_clock().now()

# Trạng thái Robot
        self.x = 0.0
        self.y = 0.0
        self.theta = 0.0

        # Trạng thái Encoder
        self.last_encoderL = None
        self.last_encoderR = None

        # Thời gian để tính dt
        self.last_time = self.get_clock().now()

    def timer_callback(self):
        # 1. Tính dt
        current_time = self.get_clock().now()
        dt = (current_time - self.last_time).nanoseconds / 1e9
        self.last_time = current_time
        if dt < 0.001:
            return
        # 2. Đọc encoder
        try:
            encoderL, encoderR = feedback()
        except Exception as e:
            self.get_logger().warn(f"Lỗi đọc Encoder: {e}")
            return

    
        #Khởi tạo lần đầu
        if self.last_encoderL is None:
            self.last_encoderL = encoderL
            self.last_encoderR = encoderR
            return

        v_l = rpm_to_mps(encoderL/18)
        v_r = rpm_to_mps(encoderR/18)
        v = (v_r - v_l) / 2.0
        omega = (v_r + v_l) / ROBOT_WHEEL_SEPARATION

        # ================== PUBLISH ODOM (VELOCITY ONLY) ==================
        odom_msg = Odometry()
        odom_msg.header.stamp = current_time.to_msg()
        odom_msg.header.frame_id = "odom"
        odom_msg.child_frame_id = "base_link"

        # ---- POSE: KHÔNG DÙNG ----
        odom_msg.pose.pose.position.x = 0.0
        odom_msg.pose.pose.position.y = 0.0
        odom_msg.pose.pose.position.z = 0.0
        odom_msg.pose.pose.orientation.w = 1.0

        odom_msg.pose.covariance = [
            0.01, 0.0, 0.0, 0.0, 0.0, 0.0,
            0.0, 0.01, 0.0, 0.0, 0.0, 0.0,
            0.0, 0.0, 99999.0, 0.0, 0.0, 0.0,
            0.0, 0.0, 0.0, 99999.0, 0.0, 0.0,
            0.0, 0.0, 0.0, 0.0, 99999.0, 0.0,
            0.0, 0.0, 0.0, 0.0, 0.0, 0.01
        ]
        odom_msg.twist.covariance = [
            0.01, 0.0, 0.0, 0.0, 0.0, 0.0,
            0.0, 0.01, 0.0, 0.0, 0.0, 0.0,
            0.0, 0.0, 99999.0, 0.0, 0.0, 0.0,
            0.0, 0.0, 0.0, 99999.0, 0.0, 0.0,
            0.0, 0.0, 0.0, 0.0, 99999.0, 0.0,
            0.0, 0.0, 0.0, 0.0, 0.0, 0.01
        ]

        # ---- TWIST: CÁI EKF DÙNG ----
        odom_msg.twist.twist.linear.x = v
        odom_msg.twist.twist.linear.y = 0.0
        odom_msg.twist.twist.angular.z = omega


        self.publisher_.publish(odom_msg)
class OdomToPath(Node):
    def __init__(self):
        super().__init__('odom_to_path')
        self.pub = self.create_publisher(Pa, '/odom_path', 10)
        self.sub = self.create_subscription(Odometry, '/odom_wheel', self.cb, 10)
        self.path = Pa()
        self.path.header.frame_id = "odom"

    def cb(self, msg: Odometry):
        pose = PoseStamped()
        pose.header = msg.header
        pose.pose = msg.pose.pose
        self.path.poses.append(pose)
        if len(self.path.poses) > 5000:
            self.path.poses.pop(0)
        self.path.header.stamp = self.get_clock().now().to_msg()
        self.pub.publish(self.path)

# ====================== CLEANUP ======================
def cleanup():
    global external_ros_processes
    print("\n🛑 Đang dọn dẹp tất cả node ROS...")

    # Kill cứng bằng tên trước (nhanh nhất)
    kill_ros_processes_by_name()

    # Kill process group nếu còn
    if external_ros_processes:
        for p in external_ros_processes:
            if p.poll() is None:
                try:
                    os.killpg(os.getpgid(p.pid), signal.SIGKILL)
                except Exception:
                    pass
        external_ros_processes.clear()

    if rclpy.ok():
        rclpy.shutdown()

    print("✅ Đã dừng sạch tất cả!")

def signal_handler(sig, frame):
    print("\n🛑 Nhận Ctrl+C – đang thoát...")
    cleanup()
    sys.exit(0)

# ====================== MAIN ======================
def main():
    global external_ros_processes

    # Launch các node bên ngoài
    launch_external_nodes()
    print("Đã khởi động các node ROS external.")

    # Khởi tạo ROS
    rclpy.init()

    # Tạo các node nội bộ
    executor = MultiThreadedExecutor()
    cmd_vel_node = CmdVelListener()
    odom_node = OdometryPublisher()
    path_node = OdomToPath()

    executor.add_node(cmd_vel_node)
    executor.add_node(odom_node)
    executor.add_node(path_node)

    # Đăng ký signal để Ctrl+C dừng sạch
    signal.signal(signal.SIGINT, signal_handler)

    print("Robot sẵn sàng! Dùng /cmd_vel để điều khiển từ RViz hoặc code khác.")
    print("Nhấn Ctrl+C để dừng toàn bộ.")

    try:
        executor.spin()
    except KeyboardInterrupt:
        pass
    finally:
        cleanup()

if __name__ == '__main__':
    main()
