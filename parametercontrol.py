# -*- coding: utf-8 -*-
"""
- 幅度顺序单参优化，支持 a1..a7 全部参与；默认顺序 a2→a1→a3→a4→a5→a6→a7（可自行改 amp_order_names）
- 相位与目标相位均为“度”（degree），Δ相位计算与包裹均在度域；相位优化可选，步长同为 (0.1, 0.01) 度
- r_k 定义：r_k = |S(k1)| / |S21|，因此分母侧出现 |S21| 时以 1 表示
- CSV 一行包含“幅度块+(相同数量)相位块”（若前置频率列会被忽略）
- 指标定义（线性幅度求和后取比）：
  a2 = (|S41|+|S51|) / (|S31|+|S21|) = (r4+r5)/(1+r3)
  a1 = (|S61|+|S71|+|S81|+|S91|) / (|S31|+|S21|+|S41|+|S51|) = (r6+r7+r8+r9)/(1+r3+r4+r5)
  a3 = (|S81|+|S91|) / (|S61|+|S71|) = (r8+r9)/(r6+r7)
  a4 = r3/1
  a5 = r5/r4
  a6 = r7/r6
  a7 = r9/r8
"""

import os, csv, math, time
from win32com.client import Dispatch
import pywintypes

GRID = 0.01  # 统一网格粒度（mm / deg）

def _fmt_list(vals, nd=6):
    try:
        return "[" + ", ".join(f"{float(v):.{nd}f}" for v in vals) + "]"
    except Exception:
        return str(vals)

def wrap_phase_deg(x):
    """角度包裹到 (-180, 180]。"""
    y = (float(x) + 180.0) % 360.0
    if y <= 0:
        y += 360.0
    return y - 180.0

def quantize_grid(x, grid=GRID):
    """对齐到网格（默认 0.01）。"""
    return round(float(x) / grid) * grid

class HFSSClosedLoopCSV:
    def __init__(self, project_path, project_name, design_name,
                 phase_param_names=None,
                 geom_param_names=None):
        print("=== 初始化 HFSS COM 接口 ===")
        self.app     = Dispatch('AnsoftHfss.HfssScriptInterface')
        self.desktop = self.app.GetAppDesktop()
        try:
            self.desktop.RestoreWindow(); print("[OK] 恢复主窗口")
        except Exception:
            print("[WARN] 无法恢复窗口")

        fullpath = os.path.join(project_path, project_name + '.aedt')
        print(f"[INFO] 打开工程：{fullpath}")
        self.project = self.desktop.OpenProject(fullpath)
        print(f"[INFO] 激活设计：{design_name}")
        self.design  = self.project.SetActiveDesign(design_name)
        print("[OK] 初始化完成。\n")

        self.p_names = phase_param_names or [f"p{i}" for i in range(1, 8)]  # p1..p7
        self.a_names = geom_param_names  or [f"a{i}" for i in range(1, 8)]  # a1..a7
        self.Nsrc = len(self.p_names)  # 7，对应端口3..9

    # ===== CSV =====
    @staticmethod
    def _parse_last_row_numeric(csv_path):
        with open(csv_path, newline='') as f:
            rd = list(csv.reader(f))
            if len(rd) < 2:
                raise RuntimeError("CSV无数据行")
            row = rd[-1]
        nums = []
        for x in row:
            try:
                nums.append(float(x))
            except Exception:
                pass
        if len(nums) < 4:
            raise RuntimeError("CSV末行可解析数值过少，请检查报表。")
        return nums

    def export_csv(self, report_name, csv_path):
        os.makedirs(os.path.dirname(csv_path), exist_ok=True)
        rep = self.design.GetModule("ReportSetup")
        rep.ExportToFile(report_name, csv_path)

    def analyze_and_get_blocks(self, setup_name, report_name, csv_path, expect_src_count):
        self.design.Analyze(setup_name)
        time.sleep(0.2)
        self.export_csv(report_name, csv_path)
        nums = self._parse_last_row_numeric(csv_path)
        total_needed = 2 * (expect_src_count + 1)  # (Nsrc+1) 幅度 + (Nsrc+1) 相位
        if len(nums) < total_needed:
            raise RuntimeError(f"CSV列数不足：解析到 {len(nums)} 个，期望 ≥ {total_needed}。")
        tail = nums[-total_needed:]
        half = len(tail) // 2
        mags   = tail[:half]
        phases = tail[half:]  # “度”
        print("    [SIM] mags    =", _fmt_list(mags, nd=6), "(linear)")
        print("    [SIM] phases° =", _fmt_list(phases, nd=6), "(degree)")
        return mags, phases

    def analyze_and_get_mags(self, setup_name, report_name, csv_path):
        mags, _ = self.analyze_and_get_blocks(setup_name, report_name, csv_path, self.Nsrc)
        return mags

    def analyze_and_get_phases_deg(self, setup_name, report_name, csv_path):
        """相位为度（degree），不转弧度。"""
        _, phases = self.analyze_and_get_blocks(setup_name, report_name, csv_path, self.Nsrc)
        return phases

    # ===== 变量写回 =====
    def _get_existing_vars(self):
        dset = set(); pset = set()
        try: dset = set(self.design.GetVariables())
        except Exception: pass
        try: pset = set(self.project.GetVariables())
        except Exception: pass
        return dset, pset

    def update_vars(self, p_vals, a_vals, do_save=True, verify=True):
        if len(p_vals) != self.Nsrc:
            raise ValueError(f"p_vals 长度应为 {self.Nsrc} (p1..p{self.Nsrc})，当前 {len(p_vals)}")
        if len(a_vals) != len(self.a_names):
            raise ValueError(f"a_vals 长度应为 {len(self.a_names)} (a1..a{len(self.a_names)})，当前 {len(a_vals)}")
        import math as _m
        flat = [*p_vals, *a_vals]
        if not all(_m.isfinite(float(x)) for x in flat):
            raise ValueError("写回变量存在 NaN/Inf")

        p_vals = [quantize_grid(v) for v in p_vals]  # 相位（度）
        a_vals = [quantize_grid(v) for v in a_vals]  # 几何（mm）

        def make_entry(name, value_str):
            return ["NAME:"+name, "PropType:=", "VariableProp",
                    "UserDef:=", True, "Value:=", value_str]

        design_vars, project_vars = self._get_existing_vars()

        def build_payload(scope_tab, server):
            tabs = ["NAME:AllTabs",
                    ["NAME:"+scope_tab,
                     ["NAME:PropServers", server]]]
            changed, new = [], []

            for name, val in zip(self.p_names, p_vals):
                entry = make_entry(name, f"{float(val)}")  # 度，写无单位
                (changed if name in (design_vars if server=="LocalVariables" else project_vars) else new).append(entry)

            for name, val in zip(self.a_names, a_vals):
                entry = make_entry(name, f"{float(val)}mm")  # mm
                (changed if name in (design_vars if server=="LocalVariables" else project_vars) else new).append(entry)

            if changed: tabs[1].append(["NAME:ChangedProps"] + changed)
            if new:     tabs[1].append(["NAME:NewProps"] + new)
            return tabs

        try:
            payload = build_payload("LocalVariableTab", "LocalVariables")
            self.design.ChangeProperty(payload)
        except pywintypes.com_error as e1:
            print("[WARN] 写设计级变量失败，改写工程级。", e1)
            payload = build_payload("ProjectVariableTab", "ProjectVariables")
            self.design.ChangeProperty(payload)

        if do_save:
            try:
                self.project.Save()
                print("[SAVE] 工程已保存")
            except Exception as e:
                print("[WARN] 保存工程失败：", e)

        print("[UPDATE] 写回：")
        print("    p =", ", ".join(f"{n}={float(v):.6f}" for n, v in zip(self.p_names, p_vals)))
        print("    a =", ", ".join(f"{n}={float(v):.6f}mm" for n, v in zip(self.a_names, a_vals)))

        if verify:
            try:
                rb_p = [self.design.GetVariableValue(n) for n in self.p_names]
                rb_a = [self.design.GetVariableValue(n) for n in self.a_names]
                print("[VERIFY] 读回：")
                print("    p' =", ", ".join(f"{n}={v}" for n, v in zip(self.p_names, rb_p)))
                print("    a' =", ", ".join(f"{n}={v}" for n, v in zip(self.a_names, rb_a)))
            except Exception as e:
                print("[WARN] 读回校验失败：", e)

    # ===== 幅度评估 r_k =====
    def eval_amp_per_port(self, mags, target_ratios=None, verbose_tag=None):
        if len(mags) != (self.Nsrc + 1):
            raise RuntimeError(f"幅度数据长度 {len(mags)} 与期望 {self.Nsrc+1} 不符。")
        ref = max(mags[0], 1e-12)            # |S21|
        ratio = [m / ref for m in mags[1:]]  # r3..r(2+Nsrc)
        if verbose_tag:
            print(f"[{verbose_tag}] ratio  = {_fmt_list(ratio)}")
        if target_ratios is None:
            return ratio, None
        if len(target_ratios) != self.Nsrc:
            raise ValueError(f"target_ratios 长度应为 {self.Nsrc}，当前 {len(target_ratios)}")
        err   = [target_ratios[i] - ratio[i] for i in range(self.Nsrc)]
        if verbose_tag:
            print(f"[{verbose_tag}] target = {_fmt_list(target_ratios)}")
            print(f"[{verbose_tag}] err    = {_fmt_list(err)}")
        return ratio, err

    @staticmethod
    def err_cost_L2(err):
        return sum(e*e for e in err)

    # ===== 指标定义（7 端口；a1..a7 全部可用） =====
    # ratios: r3..r9 => ratios[0]=r3, ..., ratios[6]=r9
    def _r(self, ratios, k):     # k = 3..9
        idx = k - 3
        if not (0 <= idx < len(ratios)):
            raise ValueError(f"需要 r{k}，但仅有 r3..r{len(ratios)+2}")
        return ratios[idx]

    def metric_value_from_ratios(self, idx, ratios, eps=1e-12):
        """
        a2 = (r4+r5)/(1+r3)
        a1 = (r6+r7+r8+r9)/(1+r3+r4+r5)
        a3 = (r8+r9)/(r6+r7)
        a4 = r3/1
        a5 = r5/r4
        a6 = r7/r6
        a7 = r9/r8
        """
        i = idx + 1  # a_i
        if i == 2:  # a2
            val = (self._r(ratios,4) + self._r(ratios,5)) / max(1.0 + self._r(ratios,3), eps)
            label = "M(a2)=(r4+r5)/(1+r3)"
        elif i == 1:  # a1
            num = sum(self._r(ratios,k) for k in range(6,10))
            den = 1.0 + sum(self._r(ratios,k) for k in range(3,6))
            val = num / max(den, eps)
            label = "M(a1)=(r6+r7+r8+r9)/(1+r3+r4+r5)"
        elif i == 3:  # a3
            val = (self._r(ratios,8) + self._r(ratios,9)) / max(self._r(ratios,6) + self._r(ratios,7), eps)
            label = "M(a3)=(r8+r9)/(r6+r7)"
        elif i == 4:  # a4
            val = self._r(ratios,3) / 1.0
            label = "M(a4)=r3/1"
        elif i == 5:  # a5
            val = self._r(ratios,5) / max(self._r(ratios,4), eps)
            label = "M(a5)=r5/r4"
        elif i == 6:  # a6
            val = self._r(ratios,7) / max(self._r(ratios,6), eps)
            label = "M(a6)=r7/r6"
        elif i == 7:  # a7
            val = self._r(ratios,9) / max(self._r(ratios,8), eps)
            label = "M(a7)=r9/r8"
        else:
            raise ValueError(f"未知 a{i} 指标映射。")
        return val, label

    def metric_target_from_target_ratios(self, idx, target_ratios, eps=1e-12):
        return self.metric_value_from_ratios(idx, target_ratios, eps=eps)[0]

    # ===== 相位评估（度制） =====
    @staticmethod
    def eval_phase_per_port_deg(phases_deg, target_phases_deg, verbose_tag=None):
        """输入与目标均为度；输出包裹后的 Δ相位（度）与绝对误差（度）。"""
        ref = phases_deg[0]  # ∠S21 (deg)
        curr_raw = [phases_deg[i] - ref for i in range(1, len(phases_deg))]  # Δ∠(Sk1 - S21) (deg)
        curr_wrapped = [wrap_phase_deg(v) for v in curr_raw]
        if len(target_phases_deg) != len(curr_wrapped):
            raise ValueError(f"target_phases 长度应为 {len(curr_wrapped)}，当前 {len(target_phases_deg)}")
        err_abs = [abs(curr_wrapped[i] - target_phases_deg[i]) for i in range(len(curr_wrapped))]
        if verbose_tag:
            print(f"[{verbose_tag}] curr(wrapped) = {_fmt_list(curr_wrapped)} deg")
            print(f"[{verbose_tag}] target       = {_fmt_list(target_phases_deg)} deg")
            print(f"[{verbose_tag}] err_abs      = {_fmt_list(err_abs)} deg")
        return curr_wrapped, err_abs

    # ===== 幅度阶段 =====
    def amplitude_block_seq(self, setup_name, report_name, csv_path,
                            p_vals, a_vals, target_ratios,
                            amp_indices,                           # 0-based
                            step_amp_list=(0.1, 0.01),            # mm（仅 0.1/0.01）
                            tol_amp=0.0001,
                            tol_metric=1e-3,
                            verbose_tag="AMP_SEQ",
                            max_iter_per_step=80,
                            improve_eps=1e-10,
                            w_global_L2=0.0):
        print(f"\n-- [{verbose_tag}] 开始（顺序单参，按 a_i 指标） --")
        print(f"[{verbose_tag}] 优化顺序：", " → ".join(self.a_names[i] for i in amp_indices))

        p_vals = [quantize_grid(v) for v in p_vals]
        a_vals = [quantize_grid(v) for v in a_vals]

        self.update_vars(p_vals, a_vals)
        mags = self.analyze_and_get_mags(setup_name, report_name, csv_path)
        ratios, err_amp = self.eval_amp_per_port(mags, target_ratios, verbose_tag=f"{verbose_tag}/BASELINE")

        for idx in amp_indices:
            aname = self.a_names[idx]
            m_target = self.metric_target_from_target_ratios(idx, target_ratios)
            m_base, m_label = self.metric_value_from_ratios(idx, ratios)
            print(f"\n[{verbose_tag}] >>> 调整 {aname} （单位 mm）")
            print(f"[{verbose_tag}] 指标 {m_label}")
            print(f"[{verbose_tag}] 目标 metric* = {m_target:.6f}")
            print(f"[{verbose_tag}] 当前 max|amp_err| = {max(abs(e) for e in err_amp):.6f}  (阈值 {tol_amp})")
            print(f"[{verbose_tag}] 当前 metric = {m_base:.6f}, err_metric = {m_target - m_base:+.6f}")

            if abs(m_target - m_base) < tol_metric:
                print(f"[{verbose_tag}] 当前指标已达标（|metric_err|<{tol_metric}），跳过 {aname}")
                continue

            for step in step_amp_list:  # 0.1 → 0.01 mm
                print(f"[{verbose_tag}] 步长 step = {step:.6f} mm")

                for it in range(1, max_iter_per_step+1):
                    base_a_vec = a_vals[:]
                    base_val   = float(base_a_vec[idx])

                    self.update_vars(p_vals, base_a_vec)
                    mags_b = self.analyze_and_get_mags(setup_name, report_name, csv_path)
                    ratios_b, err_b = self.eval_amp_per_port(mags_b, target_ratios, verbose_tag=f"{verbose_tag}/EVAL(base)")
                    m_b, _ = self.metric_value_from_ratios(idx, ratios_b)
                    cost_b = (m_b - m_target)**2 + w_global_L2 * self.err_cost_L2(err_b)
                    print(f"    iter={it:02d}: 基线 {aname}={base_val:+.6f} mm, metric={m_b:.6f}, "
                          f"metric_err={m_target - m_b:+.6f}, cost={cost_b:.6e}, max|err|={max(abs(e) for e in err_b):.6f}")

                    a_plus = base_a_vec[:];  a_plus[idx]  = quantize_grid(base_val + step)
                    self.update_vars(p_vals, a_plus)
                    mags_p = self.analyze_and_get_mags(setup_name, report_name, csv_path)
                    ratios_p, err_p = self.eval_amp_per_port(mags_p, target_ratios, verbose_tag=f"{verbose_tag}/EVAL(+)")
                    m_p, _ = self.metric_value_from_ratios(idx, ratios_p)
                    cost_p = (m_p - m_target)**2 + w_global_L2 * self.err_cost_L2(err_p)

                    a_minus = base_a_vec[:]; a_minus[idx] = quantize_grid(base_val - step)
                    self.update_vars(p_vals, a_minus)
                    mags_m = self.analyze_and_get_mags(setup_name, report_name, csv_path)
                    ratios_m, err_m = self.eval_amp_per_port(mags_m, target_ratios, verbose_tag=f"{verbose_tag}/EVAL(-)")
                    m_m, _ = self.metric_value_from_ratios(idx, ratios_m)
                    cost_m = (m_m - m_target)**2 + w_global_L2 * self.err_cost_L2(err_m)

                    print(f"      [metric] base={m_b:.6f}, +={m_p:.6f}, -={m_m:.6f}  "
                          f"[|err|_max] base={max(abs(e) for e in err_b):.6f}, +={max(abs(e) for e in err_p):.6f}, -={max(abs(e) for e in err_m):.6f}")

                    best_label, best_cost = "base", cost_b
                    best_a, best_ratios, best_err, best_metric = base_a_vec, ratios_b, err_b, m_b
                    if cost_p < best_cost - improve_eps:
                        best_label, best_cost = "plus", cost_p
                        best_a, best_ratios, best_err, best_metric = a_plus, ratios_p, err_p, m_p
                    if cost_m < best_cost - improve_eps:
                        best_label, best_cost = "minus", cost_m
                        best_a, best_ratios, best_err, best_metric = a_minus, ratios_m, err_m, m_m

                    if best_label == "base":
                        print("      → 该步未找到改进。回到基线并缩小步长。")
                        self.update_vars(p_vals, base_a_vec)
                        break
                    else:
                        a_vals  = best_a
                        ratios  = best_ratios
                        err_amp = best_err
                        print(f"      → 采纳 {best_label}：cost {cost_b:.6e} → {best_cost:.6e}, "
                              f"metric → {best_metric:.6f}, max|err| → {max(abs(e) for e in err_amp):.6f}")

                        if abs(m_target - best_metric) < tol_metric:
                            print(f"      ✔ 指标达标（|metric_err|<{tol_metric}），结束 {aname}")
                            break
                        if max(abs(e) for e in err_amp) < tol_amp:
                            print(f"      ✔ 全局幅度达标（max|err|<{tol_amp}），结束 {aname}")
                            break

                print(f"[{verbose_tag}] 完成 {aname} 当前阶段。max|err| = {max(abs(e) for e in err_amp):.6f}")

        all_ok = (max(abs(e) for e in err_amp) < tol_amp)
        print(f"\n[{verbose_tag}] 结束。最终每端口幅度误差 = {_fmt_list(err_amp)}")
        return all_ok, p_vals, a_vals, err_amp

    # ===== 相位阶段（度制；可选；步长 0.1/0.01°） =====
    def phase_block_seq(self, setup_name, report_name, csv_path,
                        p_vals, a_vals, target_phases_deg,
                        step_phase_list=(0.1, 0.01),  # deg（仅 0.1/0.01）
                        tol_phase=0.0001,
                        verbose_tag="PHASE_SEQ",
                        max_iter_per_step=80,
                        improve_eps=1e-9):
        print(f"\n-- [{verbose_tag}] 开始（顺序单参：p1→p{self.Nsrc}；单位：度） --")

        p_vals = [quantize_grid(v) for v in p_vals]
        a_vals = [quantize_grid(v) for v in a_vals]

        self.update_vars(p_vals, a_vals)
        phases_deg = self.analyze_and_get_phases_deg(setup_name, report_name, csv_path)
        curr, err_phase = self.eval_phase_per_port_deg(phases_deg, target_phases_deg, verbose_tag=f"{verbose_tag}/BASELINE")

        for idx in range(self.Nsrc):
            pname = self.p_names[idx]
            src_port = idx + 3
            print(f"\n[{verbose_tag}] >>> 调整 {pname} —— Δ∠(S{src_port}1 - S21) (deg)")
            print(f"[{verbose_tag}] 目标 target = {target_phases_deg[idx]:+.6f} deg")

            if abs(err_phase[idx]) < tol_phase:
                print(f"[{verbose_tag}] 已达标：curr={curr[idx]:+.6f} deg, err={err_phase[idx]:+.6f} deg")
                continue

            for step in step_phase_list:  # 0.1 → 0.01 deg
                print(f"[{verbose_tag}] 步长 step = {step:.6f} deg")

                for it in range(1, max_iter_per_step+1):
                    base_err_abs   = abs(err_phase[idx])
                    base_phase_val = curr[idx]
                    base_p_vec     = p_vals[:]

                    print(f"    iter={it:02d}: 基线 Δ∠={base_phase_val:+.6f} deg, err_abs={base_err_abs:.6f} deg")

                    p_plus = base_p_vec[:];  p_plus[idx]  = quantize_grid(float(p_plus[idx]) + step)
                    self.update_vars(p_plus, a_vals)
                    phases_p = self.analyze_and_get_phases_deg(setup_name, report_name, csv_path)
                    curr_p, err_p = self.eval_phase_per_port_deg(phases_p, target_phases_deg, verbose_tag=f"{verbose_tag}/EVAL(+)")
                    plus_err_abs = abs(err_p[idx])

                    p_minus = base_p_vec[:]; p_minus[idx] = quantize_grid(float(p_minus[idx]) - step)
                    self.update_vars(p_minus, a_vals)
                    phases_m = self.analyze_and_get_phases_deg(setup_name, report_name, csv_path)
                    curr_m, err_m = self.eval_phase_per_port_deg(phases_m, target_phases_deg, verbose_tag=f"{verbose_tag}/EVAL(-)")
                    minus_err_abs = abs(err_m[idx])

                    best_label = "base"; best_abs = base_err_abs
                    best_p     = base_p_vec; best_curr = curr; best_err = err_phase
                    if plus_err_abs < best_abs - improve_eps:
                        best_label, best_abs, best_p, best_curr, best_err = "plus", plus_err_abs, p_plus, curr_p, err_p
                    if minus_err_abs < best_abs - improve_eps:
                        best_label, best_abs, best_p, best_curr, best_err = "minus", minus_err_abs, p_minus, curr_m, err_m

                    if best_label == "base":
                        print("      → 该步未找到改进。回到基线并缩小步长。")
                        self.update_vars(base_p_vec, a_vals)
                        break
                    else:
                        print(f"      → 采纳 {best_label}（|err|: {base_err_abs:.6f} → {best_abs:.6f}）")
                        p_vals    = best_p
                        curr      = best_curr
                        err_phase = best_err
                        if best_abs < tol_phase:
                            print(f"      ✔ {pname} 达标（|err|<{tol_phase} deg）")
                            break

            if abs(err_phase[idx]) >= tol_phase:
                print(f"[{verbose_tag}] 警告：{pname} 未达标，残差 {err_phase[idx]:+.6f} deg")

        all_ok = all(abs(e) < tol_phase for e in err_phase)
        print(f"\n[{verbose_tag}] 结束。每端口相位误差(deg) = {_fmt_list(err_phase)}")
        return all_ok, p_vals, a_vals, err_phase

    # ===== 主流程 =====
    def closed_loop_amp_then_phase(self,
                                   setup_name, report_name, csv_path,
                                   target_ratios, target_phases_deg,
                                   init_p, init_a,
                                   adjust_amplitude=True,
                                   adjust_phase=False,   # 相位可选
                                   amp_indices=None,
                                   step_amp_list=(0.1, 0.01),    # mm（仅 0.1/0.01）
                                   tol_amp=0.0001,
                                   tol_metric=1e-3,
                                   step_phase_list=(0.1, 0.01),  # deg（仅 0.1/0.01）
                                   tol_phase=0.0001,
                                   max_iter_per_step_amp=80,
                                   max_iter_per_step_phase=80,
                                   improve_eps_amp=1e-10,
                                   improve_eps_phase=1e-9,
                                   w_global_L2=0.0):
        if len(init_p) != self.Nsrc:
            raise ValueError(f"init_p 长度应为 {self.Nsrc}")
        if len(init_a) != len(self.a_names):
            raise ValueError(f"init_a 长度应为 {len(self.a_names)}")
        if len(target_ratios) != self.Nsrc:
            raise ValueError(f"target_ratios 长度应为 {self.Nsrc}")
        if len(target_phases_deg) != self.Nsrc:
            raise ValueError(f"target_phases 长度应为 {self.Nsrc}")

        p_curr, a_curr = [quantize_grid(v) for v in init_p], [quantize_grid(v) for v in init_a]
        err_amp  = None
        err_phase = None
        ok_amp = ok_phase = True

        if adjust_amplitude:
            # 若未指定，默认对 a1..a7 依次优化
            if amp_indices is None:
                amp_indices = list(range(len(self.a_names)))
            print("\n========== 阶段 A：幅度优化（0.1 → 0.01 mm；顺序按 amp_indices 指定） ==========")
            ok_amp, p_curr, a_curr, err_amp = self.amplitude_block_seq(
                setup_name, report_name, csv_path,
                p_vals=p_curr, a_vals=a_curr, target_ratios=target_ratios,
                amp_indices=amp_indices,
                step_amp_list=step_amp_list,
                tol_amp=tol_amp,
                tol_metric=tol_metric,
                verbose_tag="AMP_SEQ",
                max_iter_per_step=max_iter_per_step_amp,
                improve_eps=improve_eps_amp,
                w_global_L2=w_global_L2
            )
        else:
            print("\n[AMP_SEQ] 跳过幅度优化，仅计算基线幅度误差用于报告。")
            self.update_vars(p_curr, a_curr)
            mags0 = self.analyze_and_get_mags(setup_name, report_name, csv_path)
            _, err_amp = self.eval_amp_per_port(mags0, target_ratios, verbose_tag="AMP_SEQ/BASELINE")

        if adjust_phase:
            print("\n========== 阶段 P：相位优化（0.1 → 0.01 deg） ==========")
            ok_phase, p_curr, a_curr, err_phase = self.phase_block_seq(
                setup_name, report_name, csv_path,
                p_vals=p_curr, a_vals=a_curr, target_phases_deg=target_phases_deg,
                step_phase_list=step_phase_list,
                tol_phase=tol_phase,
                verbose_tag="PHASE_SEQ",
                max_iter_per_step=max_iter_per_step_phase,
                improve_eps=improve_eps_phase
            )
        else:
            print("\n[PHASE_SEQ] 相位保持不变（不优化）。仅报告当前相位误差（度）。")
            self.update_vars(p_curr, a_curr)
            ph0_deg = self.analyze_and_get_phases_deg(setup_name, report_name, csv_path)
            _, err_phase = self.eval_phase_per_port_deg(ph0_deg, target_phases_deg, verbose_tag="PHASE_SEQ/BASELINE")

        self.update_vars(p_curr, a_curr, do_save=True, verify=True)
        try:
            self.project.Save()
        except Exception:
            pass

        print("\n[SUMMARY] 阶段结果：")
        print("  幅度阶段启用  :", adjust_amplitude, "；达标:", ok_amp)
        print("  相位阶段启用  :", adjust_phase,   ("；达标:"+str(ok_phase)) if adjust_phase else "；已禁用（仅报告度制误差）")
        return (ok_amp and (not adjust_phase or ok_phase)), p_curr, a_curr, err_amp, err_phase


if __name__ == "__main__":
    # ====== 用户配置 ======
    project_path = r"C:\\Users\\ji0035li\\Work Folders\\Desktop"
    project_name = "SIW-Slot-Array-32 for JWl"
    design_name  = "SIW-Slot-Array-32-ELC1"

    setup_name   = "Setup1"
    report_name  = "S Parameter Plot 1"
    csv_path     = r"C:\\temp\\sparams_amp_phase.csv"

    # 7 个参数名
    phase_param_names = [f"p{i}" for i in range(1,8)]   # p1..p7（度）
    geom_param_names  = [f"a{i}" for i in range(1,8)]   # a1..a7（mm）

    # 目标（长度均为 7）
    # target_ratios 对应 r3..r9；target_phases_deg 对应 Δ∠(S31−S21)..Δ∠(S91−S21)（单位：度）
    target_ratios     = [1, 2, 2, 1.684995697, 1.684995697, 0.973018554, 0.973018554]
    target_phases_deg = [-10.3362055, -98.9599246, -133.0633558, 184.7587902, 113.5462629, 100.4418957, -139.1326191]

    # 初值（示例；相位为度，几何为 mm）
    init_p = [-1.190000, -2.470000, -2.800000, 2.640000, 1.450000, 0.940000, -3.280000]
    init_a = [ 0.120000, -0.250000, 0.510000, 0.040000, -0.420000, 0.230000, -0.290000]

    # 开关：按需启用相位优化（度制 & 步长 0.1/0.01）
    DO_AMPLITUDE = True
    DO_PHASE     = True  # 若需相位闭环，设 True

    # 幅度顺序（默认 a2 → a1 → a3 → a4 → a5 → a6 → a7）
    amp_order_names = ["a2", "a1", "a3", "a4", "a5", "a6", "a7"]
    amp_indices = [int(name[1:]) - 1 for name in amp_order_names]

    hf = HFSSClosedLoopCSV(project_path, project_name, design_name,
                           phase_param_names=phase_param_names,
                           geom_param_names=geom_param_names)

    ok, p_final, a_final, err_amp, err_phase_deg = hf.closed_loop_amp_then_phase(
        setup_name=setup_name, report_name=report_name, csv_path=csv_path,
        target_ratios=target_ratios, target_phases_deg=target_phases_deg,
        init_p=init_p, init_a=init_a,
        adjust_amplitude=DO_AMPLITUDE,
        adjust_phase=DO_PHASE,                 # 相 【位闭环开关
        amp_indices=amp_indices,
        step_amp_list=(0.1, 0.01),             # 幅度双粒度（mm）
        tol_amp=0.0001,
        tol_metric=1e-3,
        step_phase_list=(0.1, 0.01),           # 相位双粒度（deg）
        tol_phase=0.0001,
        max_iter_per_step_amp=80,
        max_iter_per_step_phase=80,
        improve_eps_amp=1e-10,
        improve_eps_phase=1e-9,
        w_global_L2=0.0
    )

    print("\n[FINAL] 全流程完成：", ok)
    print("p_final (deg)            =", _fmt_list(p_final, nd=6))
    print("a_final (mm)             =", _fmt_list(a_final, nd=6))
    print("final amp errs           =", _fmt_list(err_amp if err_amp is not None else [], nd=6))
    print("final phase errs (deg)   =", _fmt_list(err_phase_deg if err_phase_deg is not None else [], nd=6))

    # 双保险保存
    hf.update_vars(p_final, a_final, do_save=True, verify=True)
    try:
        hf.project.Save()
        print("[SAVE] 工程保存成功（final）")
    except Exception as e:
        print("[WARN] 最终保存失败：", e)
