"""
SCARA Robot 4-DOF Simulator
- Dong hoc thuan (FK): tinh toa do tu goc khop
- Dong hoc nghich (IK): tinh goc khop tu toa do dich
- GUI voi thanh truot va nhap lieu
- Hien thi 3D truc quan
"""

import tkinter as tk
from tkinter import ttk, messagebox
import numpy as np
import matplotlib
matplotlib.use('TkAgg')
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.patches import FancyArrowPatch
from mpl_toolkits.mplot3d import Axes3D
import warnings
warnings.filterwarnings('ignore')


# ─────────────── KINEAMATICS ───────────────
class SCARAKinematics:
    def __init__(self, L1=200, L2=100):
        self.L1 = L1
        self.L2 = L2

    def forward(self, theta1_deg, theta2_deg, theta3_deg, d4):
        t1 = np.radians(theta1_deg)
        t2 = np.radians(theta2_deg)

        # 1. Tính toán X, Y
        x = self.L1 * np.cos(t1) + self.L2 * np.cos(t1 + t2)
        y = self.L1 * np.sin(t1) + self.L2 * np.sin(t1 + t2)
        
        # 2. Toàn bộ cánh tay di chuyển theo d4
        # Giả sử 200 là độ cao gốc, d4 là khoảng dịch chuyển tịnh tiến
        current_z =  d4 
        
        phi = theta1_deg + theta2_deg + theta3_deg 

        # 3. Vị trí các khớp (Di chuyển đồng bộ theo current_z)
        j0 = np.array([0, 0, 0])
        
        # J1 nâng/hạ theo d4
        j1 = np.array([0, 0, current_z]) 
        
        # J2 cũng nâng/hạ theo d4 và xoay theo t1
        j2 = np.array([
            self.L1 * np.cos(t1), 
            self.L1 * np.sin(t1), 
            current_z
        ])
        
        # Điểm cuối EE cũng ở cùng độ cao này
        ee = np.array([x, y, current_z])

        return {
            'x': x, 'y': y, 'z': current_z, 'phi': phi,
            'joints': [j0, j1, j2, ee]
        }

    def inverse(self, x, y, z, phi_deg, elbow=1):
        """Dong hoc nghich: (x, y, z, phi) -> (theta1, theta2, theta3, d4)"""
        L1, L2 = self.L1, self.L2
        r_sq = x**2 + y**2
        c2 = (r_sq - L1**2 - L2**2) / (2 * L1 * L2)

        if abs(c2) > 1:
            return None, "Ngoai tam voi (out of reach)!"

        s2 = elbow * np.sqrt(max(0, 1 - c2**2))
        theta2 = np.degrees(np.arctan2(s2, c2))
        theta1 = np.degrees(np.arctan2(y, x) - np.arctan2(L2 * s2, L1 + L2 * c2))
        theta3 = phi_deg - theta1 - theta2
        d4 = z

        # Kiem tra gioi han
        if not (-170 <= theta1 <= 170):
            return None, f"theta1={theta1:.1f}° vuot gioi han [-170, 170]"
        if not (-140 <= theta2 <= 140):
            return None, f"theta2={theta2:.1f}° vuot gioi han [-140, 140]"
        if not (-180 <= theta3 <= 180):
            return None, f"theta3={theta3:.1f}° vuot gioi han [-180, 180]"
        if not (0 <= d4 <= 200):
            return None, f"d4={d4:.1f} vuot gioi han [0, 120]"

        return {'theta1': theta1, 'theta2': theta2, 'theta3': theta3, 'd4': d4}, None


# ─────────────── GUI ───────────────
class SCARAApp:
    def __init__(self, root):
        self.root = root
        self.root.title("SCARA Robot 4-DOF — Dong hoc Thuan/Nghich")
        self.root.configure(bg='#1e1e1e')

        self.kin = SCARAKinematics()
        self._anim_running = False
        self._anim_id = None
        self._anim_t = 0.0
        self._anim_dir = 1

        self._build_ui()
        self._update_fk()

    # ── Xay dung giao dien ──
    def _build_ui(self):
        root = self.root

        # ── Matplotlib Figure ──
        self.fig = plt.figure(figsize=(7, 5.5), facecolor='#1e1e1e')
        self.ax = self.fig.add_subplot(111, projection='3d', facecolor='#1e1e1e')
        self.canvas = FigureCanvasTkAgg(self.fig, master=root)
        self.canvas.get_tk_widget().pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=4, pady=4)

        # ── Panel ben phai ──
        panel = tk.Frame(root, bg='#1e1e1e', width=320)
        panel.pack(side=tk.RIGHT, fill=tk.Y, padx=4, pady=4)
        panel.pack_propagate(False)

        # Scrollable canvas bên phải
        canvas_frame = tk.Canvas(panel, bg='#1e1e1e', highlightthickness=0)
        scrollbar = ttk.Scrollbar(panel, orient='vertical', command=canvas_frame.yview)
        self.scroll_frame = tk.Frame(canvas_frame, bg='#1e1e1e')
        self.scroll_frame.bind('<Configure>', lambda e: canvas_frame.configure(scrollregion=canvas_frame.bbox('all')))
        canvas_frame.create_window((0, 0), window=self.scroll_frame, anchor='nw')
        canvas_frame.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        p = self.scroll_frame

        def section(text):
            f = tk.LabelFrame(p, text=text, bg='#2a2a2a', fg='#9fe1cb',
                              font=('Consolas', 10, 'bold'), bd=1, relief='groove',
                              padx=8, pady=6)
            f.pack(fill=tk.X, padx=4, pady=4)
            return f

        # ── Thong so robot ──
        frm_params = section("⚙  Thong so robot")
        self.L1_var = tk.DoubleVar(value=200)
        self.L2_var = tk.DoubleVar(value=100)
        self._param_slider(frm_params, "L₁ (mm)", self.L1_var, 60, 220, self._on_param_change)
        self._param_slider(frm_params, "L₂ (mm)", self.L2_var, 50, 180, self._on_param_change)

        # ── Dong hoc thuan ──
        frm_fk = section("🔵  Dong hoc Thuan (FK)")
        self.t1_var = tk.DoubleVar(value=30)
        self.t2_var = tk.DoubleVar(value=-45)
        self.t3_var = tk.DoubleVar(value=20)
        self.d4_var = tk.DoubleVar(value=40)
        self._joint_slider(frm_fk, "θ₁ (°)", self.t1_var, -170, 170, self._update_fk)
        self._joint_slider(frm_fk, "θ₂ (°)", self.t2_var, -140, 140, self._update_fk)
        self._joint_slider(frm_fk, "θ₃ (°)", self.t3_var, -180, 180, self._update_fk)
        self._joint_slider(frm_fk, "d₄ (mm)", self.d4_var, 0, 200, self._update_fk)

        # Ket qua FK
        frm_fk_res = section("📍 Ket qua FK — Vi tri TCP")
        self.lbl_x = self._result_row(frm_fk_res, "X (mm)")
        self.lbl_y = self._result_row(frm_fk_res, "Y (mm)")
        self.lbl_z = self._result_row(frm_fk_res, "Z (mm)")
        self.lbl_phi = self._result_row(frm_fk_res, "φ (°)")

        # ── Dong hoc nghich ──
        frm_ik = section("🟠  Dong hoc Nghich (IK)")

        def ik_entry(parent, label, default):
            row = tk.Frame(parent, bg='#2a2a2a')
            row.pack(fill=tk.X, pady=2)
            tk.Label(row, text=label, bg='#2a2a2a', fg='#c2c0b6',
                     font=('Consolas', 9), width=8, anchor='w').pack(side=tk.LEFT)
            var = tk.DoubleVar(value=default)
            e = tk.Entry(row, textvariable=var, bg='#333', fg='#e8e5dc',
                         font=('Consolas', 10), width=10, bd=0,
                         insertbackground='white', relief='flat',
                         highlightthickness=1, highlightbackground='#555')
            e.pack(side=tk.LEFT, padx=4)
            return var

        self.ik_x = ik_entry(frm_ik, "X (mm)", 200)
        self.ik_y = ik_entry(frm_ik, "Y (mm)", 150)
        self.ik_z = ik_entry(frm_ik, "Z (mm)", 160)
        self.ik_phi = ik_entry(frm_ik, "φ (°)", 0)

        btn_row = tk.Frame(frm_ik, bg='#2a2a2a')
        btn_row.pack(fill=tk.X, pady=6)
        tk.Button(btn_row, text="Nghiem 1 (elbow+)", command=lambda: self._solve_ik(1),
                  bg='#185FA5', fg='white', font=('Consolas', 9, 'bold'),
                  relief='flat', padx=8, pady=4, cursor='hand2').pack(side=tk.LEFT, padx=(0, 4))
        tk.Button(btn_row, text="Nghiem 2 (elbow-)", command=lambda: self._solve_ik(-1),
                  bg='#3B6D11', fg='white', font=('Consolas', 9, 'bold'),
                  relief='flat', padx=8, pady=4, cursor='hand2').pack(side=tk.LEFT)

        self.lbl_ik_status = tk.Label(frm_ik, text="", bg='#2a2a2a',
                                       font=('Consolas', 9), wraplength=280, justify='left')
        self.lbl_ik_status.pack(fill=tk.X, pady=2)

        frm_ik_res = section("📐 Ket qua IK — Goc khop")
        self.lbl_iq1 = self._result_row(frm_ik_res, "θ₁ (°)")
        self.lbl_iq2 = self._result_row(frm_ik_res, "θ₂ (°)")
        self.lbl_iq3 = self._result_row(frm_ik_res, "θ₃ (°)")
        self.lbl_iq4 = self._result_row(frm_ik_res, "d₄")

        # ── Nut dieu khien ──
        frm_ctrl = section("🎮  Dieu khien")
        ctrl_row = tk.Frame(frm_ctrl, bg='#2a2a2a')
        ctrl_row.pack(fill=tk.X)
        tk.Button(ctrl_row, text="Reset", command=self._reset,
                  bg='#5F5E5A', fg='white', font=('Consolas', 9, 'bold'),
                  relief='flat', padx=10, pady=4, cursor='hand2').pack(side=tk.LEFT, padx=(0, 6))
        self.btn_demo = tk.Button(ctrl_row, text="▶ Demo", command=self._toggle_demo,
                  bg='#993C1D', fg='white', font=('Consolas', 9, 'bold'),
                  relief='flat', padx=10, pady=4, cursor='hand2')
        self.btn_demo.pack(side=tk.LEFT, padx=(0, 6))
        tk.Button(ctrl_row, text="Luu hinh", command=self._save_fig,
                  bg='#3C3489', fg='white', font=('Consolas', 9, 'bold'),
                  relief='flat', padx=10, pady=4, cursor='hand2').pack(side=tk.LEFT)

    # ── Helper widgets ──
    def _param_slider(self, parent, label, var, lo, hi, cmd):
        row = tk.Frame(parent, bg='#2a2a2a')
        row.pack(fill=tk.X, pady=2)
        tk.Label(row, text=label, bg='#2a2a2a', fg='#c2c0b6',
                 font=('Consolas', 9), width=10, anchor='w').pack(side=tk.LEFT)
        val_lbl = tk.Label(row, textvariable=var, bg='#2a2a2a', fg='#9fe1cb',
                           font=('Consolas', 9), width=5)
        val_lbl.pack(side=tk.RIGHT)
        s = tk.Scale(row, from_=lo, to=hi, orient=tk.HORIZONTAL, variable=var,
                     command=lambda v: cmd(),
                     bg='#2a2a2a', fg='#c2c0b6', troughcolor='#444',
                     highlightthickness=0, bd=0, sliderrelief='flat',
                     length=140, resolution=1, showvalue=False)
        s.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=4)

    def _joint_slider(self, parent, label, var, lo, hi, cmd):
        row = tk.Frame(parent, bg='#2a2a2a')
        row.pack(fill=tk.X, pady=2)
        tk.Label(row, text=label, bg='#2a2a2a', fg='#c2c0b6',
                 font=('Consolas', 9), width=9, anchor='w').pack(side=tk.LEFT)
        val_lbl = tk.Label(row, textvariable=var, bg='#2a2a2a', fg='#FAC775',
                           font=('Consolas', 9), width=6)
        val_lbl.pack(side=tk.RIGHT)
        s = tk.Scale(row, from_=lo, to=hi, orient=tk.HORIZONTAL, variable=var,
                     command=lambda v: cmd(),
                     bg='#2a2a2a', fg='#c2c0b6', troughcolor='#444',
                     highlightthickness=0, bd=0, sliderrelief='flat',
                     length=130, resolution=1, showvalue=False)
        s.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=4)

    def _result_row(self, parent, label):
        row = tk.Frame(parent, bg='#2a2a2a')
        row.pack(fill=tk.X, pady=1)
        tk.Label(row, text=label, bg='#2a2a2a', fg='#888780',
                 font=('Consolas', 9), width=9, anchor='w').pack(side=tk.LEFT)
        lbl = tk.Label(row, text="—", bg='#2a2a2a', fg='#e8e5dc',
                       font=('Consolas', 10, 'bold'))
        lbl.pack(side=tk.LEFT)
        return lbl

    # ── Cap nhat FK ──
    def _update_fk(self, *_):
        self.kin.L1 = self.L1_var.get()
        self.kin.L2 = self.L2_var.get()
        res = self.kin.forward(self.t1_var.get(), self.t2_var.get(),
                               self.t3_var.get(), self.d4_var.get())
        self.lbl_x.config(text=f"{res['x']:.2f}")
        self.lbl_y.config(text=f"{res['y']:.2f}")
        self.lbl_z.config(text=f"{res['z']:.2f}")
        self.lbl_phi.config(text=f"{res['phi']:.2f}°")
        self._draw(res)

    def _on_param_change(self):
        self._update_fk()

    # ── Giai IK ──
    def _solve_ik(self, elbow):
        try:
            X = self.ik_x.get()
            Y = self.ik_y.get()
            Z = self.ik_z.get()
            phi = self.ik_phi.get()
        except Exception:
            messagebox.showerror("Loi", "Vui long nhap so hop le!")
            return

        sol, err = self.kin.inverse(X, Y, Z, phi, elbow)
        if err:
            self.lbl_ik_status.config(text=f"❌ {err}", fg='#F09595')
            for lbl in [self.lbl_iq1, self.lbl_iq2, self.lbl_iq3, self.lbl_iq4]:
                lbl.config(text="—")
            return

        self.lbl_ik_status.config(text="✅ Giai thanh cong!", fg='#5DCAA5')
        self.lbl_iq1.config(text=f"{sol['theta1']:.2f}°")
        self.lbl_iq2.config(text=f"{sol['theta2']:.2f}°")
        self.lbl_iq3.config(text=f"{sol['theta3']:.2f}°")
        self.lbl_iq4.config(text=f"{sol['d4']:.2f}")

        # Cap nhat slider sang vi tri IK
        self.t1_var.set(round(sol['theta1'], 1))
        self.t2_var.set(round(sol['theta2'], 1))
        self.t3_var.set(round(sol['theta3'], 1))
        self.d4_var.set(round(sol['d4'], 1))
        self._update_fk()

    # ── Ve robot 3D ──
    def _draw(self, res):
        ax = self.ax
        ax.cla()
        ax.set_facecolor('#1e1e1e')
        self.fig.patch.set_facecolor('#1e1e1e')

        joints = res['joints']  # j0, j1, j2, ee
        j0, j1, j2, ee = joints

        L1, L2 = self.kin.L1, self.kin.L2
        maxR = L1 + L2

        # ── Nen (ground plane) ──
        u = np.linspace(-maxR * 1.1, maxR * 1.1, 2)
        v = np.linspace(-maxR * 1.1, maxR * 1.1, 2)
        U, V = np.meshgrid(u, v)
        ax.plot_surface(U, V, np.zeros_like(U), alpha=0.08, color='#4a4a44', linewidth=0)

        # Grid
        for gv in np.arange(-maxR, maxR + 1, 60):
            ax.plot([gv, gv], [-maxR, maxR], [0, 0], color='#3a3a38', linewidth=0.4)
            ax.plot([-maxR, maxR], [gv, gv], [0, 0], color='#3a3a38', linewidth=0.4)

        # ── Vung tam voi (workspace) ──
        angles = np.linspace(0, 2 * np.pi, 120)
        d4_cur = self.d4_var.get()
        ax.plot(maxR * np.cos(angles), maxR * np.sin(angles),
                np.full(120, d4_cur), color='#185FA5', linewidth=0.8, alpha=0.5, linestyle='--')
        minR = abs(L1 - L2)
        if minR > 5:
            ax.plot(minR * np.cos(angles), minR * np.sin(angles),
                np.full(120, d4_cur), color='#185FA5', linewidth=0.5, alpha=0.3, linestyle=':')

        # ── Truc toa do ──
        orig = np.array([0, 0, 0])
        for vec, col, lbl in [([80, 0, 0], '#E24B4A', 'X'),
                               ([0, 80, 0], '#639922', 'Y'),
                               ([0, 0, 80], '#378ADD', 'Z')]:
            ax.quiver(*orig, *vec, color=col, linewidth=1.2, arrow_length_ratio=0.2)
            ax.text(*(orig + np.array(vec) * 1.15), lbl, color=col, fontsize=8, fontweight='bold')

        # ── De robot ──
        theta_base = np.linspace(0, 2 * np.pi, 40)
        R_base = 32
        for h in np.linspace(0, self.kin.L1, 3):
            ax.plot(R_base * np.cos(theta_base), R_base * np.sin(theta_base),
                    np.full(40, h), color='#888780', linewidth=0.7, alpha=0.5)

        # ── Lien ket 1 (xanh duong) ──
        # ax.plot([j0[0], j1[0]], [j0[1], j1[1]], [j0[2], j1[2]],
        #         color='#378ADD', linewidth=8, solid_capstyle='round',
        #         label=f'Lien ket 1 (L₁={int(L1)}mm)')

        L1 = self.kin.L1
        ax.plot([0, 0], [0, 0], [0, L1],
                color='#888780', linewidth=8, solid_capstyle='round',
                label=f'Truc L1 co dinh ({int(L1)}mm)')

        # Rồi mới vẽ con trượt d4 (chấm tròn) tại vị trí j1
        ax.scatter(*j1, s=200, c='#378ADD', zorder=10, depthshade=False)

        # ── Lien ket 2 (xanh la) ──
        ax.plot([j1[0], j2[0]], [j1[1], j2[1]], [j1[2], j2[2]],
                color='#1D9E75', linewidth=6, solid_capstyle='round',
                label=f'Lien ket 2 (L₂={int(L2)}mm)')

        # ── Khop prismatic (tim) ──
        ax.plot([j2[0], ee[0]], [j2[1], ee[1]], [j2[2], ee[2]],
                color='#7F77DD', linewidth=5, solid_capstyle='round',
                label=f'd₄={int(self.d4_var.get())}mm')

        # ── Shadow tren san ──
        ax.plot([j0[0], j1[0], j2[0]], [j0[1], j1[1], j2[1]],
                [0, 0, 0], color='#333', linewidth=3, alpha=0.5)

        # ── Ve khop ──
        for pos, size, col, lbl in [
            (j0, 80, '#5F5E5A', 'J0'),
            (j1, 60, '#185FA5', 'J1'),
            (j2, 50, '#0F6E56', 'J2'),
        ]:
            u2 = np.linspace(0, 2 * np.pi, 30)
            v2 = np.linspace(0, np.pi, 20)
            xs = (size / 10) * np.outer(np.cos(u2), np.sin(v2)) + pos[0]
            ys = (size / 10) * np.outer(np.sin(u2), np.sin(v2)) + pos[1]
            zs = (size / 10) * np.outer(np.ones_like(u2), np.cos(v2)) + pos[2]
            ax.plot_surface(xs, ys, zs, color=col, alpha=0.9, linewidth=0)
            ax.text(pos[0], pos[1], pos[2] + size / 8, lbl,
                    color='white', fontsize=7, ha='center', va='bottom', fontweight='bold')

        # ── End-effector ──
        ax.scatter(*ee, s=120, c='#D85A30', zorder=10, depthshade=False)
        # Huong EE
        phi_rad = np.radians(res['phi'])
        ee_len = 28
        ee_dir = np.array([ee_len * np.cos(phi_rad), ee_len * np.sin(phi_rad), 0])
        ax.quiver(*ee, *ee_dir, color='#D85A30', linewidth=2,
                  arrow_length_ratio=0.35, label='End-effector')

        # ── Kich thuoc chu thich ──
        ax.text(ee[0] + 6, ee[1] + 6, ee[2] + 10,
                f"TCP\n({res['x']:.0f}, {res['y']:.0f}, {res['z']:.0f})",
                color='#FAC775', fontsize=7.5, fontweight='bold',
                bbox=dict(boxstyle='round,pad=0.3', facecolor='#1e1e1e',
                          edgecolor='#FAC775', alpha=0.8))

        # ── Format truc ──
        ax.set_xlim(-maxR * 1.2, maxR * 1.2)
        ax.set_ylim(-maxR * 1.2, maxR * 1.2)
        ax.set_zlim(0, 260)
        ax.set_xlabel('X', color='#E24B4A', fontsize=9, labelpad=4)
        ax.set_ylabel('Y', color='#639922', fontsize=9, labelpad=4)
        ax.set_zlabel('Z', color='#378ADD', fontsize=9, labelpad=4)
        ax.tick_params(colors='#888780', labelsize=7)
        ax.xaxis.pane.fill = False
        ax.yaxis.pane.fill = False
        ax.zaxis.pane.fill = False
        ax.xaxis.pane.set_edgecolor('#2a2a2a')
        ax.yaxis.pane.set_edgecolor('#2a2a2a')
        ax.zaxis.pane.set_edgecolor('#2a2a2a')
        ax.grid(True, color='#2a2a2a', linewidth=0.5)

        leg = ax.legend(loc='upper left', fontsize=7, framealpha=0.6,
                        facecolor='#1e1e1e', edgecolor='#444', labelcolor='#c2c0b6')

        self.fig.tight_layout(pad=0.5)
        self.canvas.draw_idle()

    # ── Demo animation ──
    def _toggle_demo(self):
        if self._anim_running:
            self._anim_running = False
            self.btn_demo.config(text="▶ Demo")
        else:
            self._anim_running = True
            self.btn_demo.config(text="⏹ Dung")
            self._anim_step()

    def _anim_step(self):
        if not self._anim_running:
            return
        self._anim_t += 0.015 * self._anim_dir
        if self._anim_t >= 1 or self._anim_t <= 0:
            self._anim_dir *= -1
        t = self._anim_t
        self.t1_var.set(round(-80 + 160 * t, 1))
        self.t2_var.set(round(-120 + 240 * t, 1))
        self.t3_var.set(round(-60 + 120 * t, 1))
        self.d4_var.set(round(5 + 90 * t, 1))
        self._update_fk()
        self._anim_id = self.root.after(40, self._anim_step)

    # ── Reset ──
    def _reset(self):
        self._anim_running = False
        self.btn_demo.config(text="▶ Demo")
        self.t1_var.set(30)
        self.t2_var.set(-45)
        self.t3_var.set(20)
        self.d4_var.set(40)
        self.L1_var.set(200)
        self.L2_var.set(100)
        self._update_fk()

    def _save_fig(self):
        self.fig.savefig('scara_robot.png', dpi=150, bbox_inches='tight',
                         facecolor='#1e1e1e')
        messagebox.showinfo("Da luu", "Hinh da duoc luu vao: scara_robot.png")


# ─────────────── MAIN ───────────────
if __name__ == '__main__':
    root = tk.Tk()
    root.geometry('1100x680')
    root.minsize(900, 580)

    style = ttk.Style()
    style.theme_use('clam')
    style.configure('TScrollbar', troughcolor='#2a2a2a', background='#555', arrowcolor='#888')

    app = SCARAApp(root)
    root.mainloop()