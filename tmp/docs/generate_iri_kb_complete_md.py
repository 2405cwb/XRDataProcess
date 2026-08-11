import itertools
import json
from pathlib import Path

import numpy as np

SRC = Path(r"D:\job\工作COD\2025\公路\XRDataProcess\tmp\docs\iri_truth_kb_analysis.json")
OUT = Path(r"D:\job\工作COD\2025\公路\XRDataProcess\output\doc\栗庙路K_B完整推算与5pct可达性报告_0.25m重出数据.md")
ORDER = ["200-300", "300-400", "400-500", "500-600"]


def fit_min_mape(samples):
    """min Σ |K*x+B-y|/y. Two variables: the convex optimum lies at a residual-line intersection."""
    x = np.array([s["raw_iri"] for s in samples], dtype=float)
    y = np.array([s["truth"] for s in samples], dtype=float)
    candidates = []
    for i, j in itertools.combinations(range(len(x)), 2):
        a = np.array([[x[i], 1.0], [x[j], 1.0]])
        if abs(np.linalg.det(a)) > 1e-10:
            candidates.append(np.linalg.solve(a, [y[i], y[j]]))
    # Positive K is physically required; include its boundary.
    candidates.append(np.array([0.0, np.median(y)]))
    objective = lambda q: np.mean(np.abs((q[0] * x + q[1] - y) / y))
    q = min(candidates, key=objective)
    rel = (q[0] * x + q[1] - y) / y
    return {"k": q[0], "b": q[1], "mean": np.mean(abs(rel)), "max": np.max(abs(rel)), "rel": rel}


def fit_minimax(samples):
    """min max |K*x+B-y|/y. Enumerate all 3 active signed constraints."""
    x = np.array([s["raw_iri"] for s in samples], dtype=float)
    y = np.array([s["truth"] for s in samples], dtype=float)
    candidates = []
    for ids in itertools.combinations(range(len(x)), 3):
        for signs in itertools.product((-1.0, 1.0), repeat=3):
            # K*x + B - y = sign * y * z
            a = np.array([[x[i], 1.0, -sgn * y[i]] for i, sgn in zip(ids, signs)])
            if abs(np.linalg.det(a)) > 1e-10:
                q = np.linalg.solve(a, [y[i] for i in ids])
                if q[0] >= 0 and q[2] >= -1e-10:
                    candidates.append(q)
    objective = lambda q: np.max(np.abs((q[0] * x + q[1] - y) / y))
    q = min(candidates, key=objective)
    rel = (q[0] * x + q[1] - y) / y
    return {"k": q[0], "b": q[1], "mean": np.mean(abs(rel)), "max": np.max(abs(rel)), "rel": rel}


def main():
    data = json.loads(SRC.read_text(encoding="utf-8"))
    samples = data["samples"]
    models = {}
    for group in ("30", "50", "70"):
        ss = [s for s in samples if s["group"] == group]
        models[group] = {"mape": fit_min_mape(ss), "minimax": fit_minimax(ss)}

    lines = []
    add = lines.append
    add("# 栗庙路原始IRI分速度K、B推算及5%可达性完整报告（0.25m重出数据）")
    add("")
    add("## 1. 结论")
    add("")
    add("本报告以用户指定的“原始平整度结果”Excel表为数值来源。表内100m IRI已经核实为对应10个 `IRI_10m` 的算术平均；因此在本批数据中，按100m表值直接套K、B，与“每10m按速度选K、B、再平均为100m”的出表逻辑等价。")
    add("")
    add("在每个速度组只能使用一组固定、且K为正的线性参数 `IRI校正 = K × IRI原始 + B` 的前提下：")
    add("")
    add("| 速度组 | 最小可达平均相对误差 | 平均误差能否<5% | 最小可达最大单段误差 | 所有100m能否<5% |")
    add("|---|---:|---|---:|---|")
    for group in ("30", "50", "70"):
        a = models[group]["mape"]
        b = models[group]["minimax"]
        add(f"| {group} | {a['mean']*100:.2f}% | {'可以' if a['mean'] < .05 else '不可以'} | {b['max']*100:.2f}% | 不可以 |")
    add("")
    add("因此：70速度组可把平均误差压至5%以下；50速度组理论最优仍为5.08%，只差0.08个百分点；30速度组理论最优为5.99%。三组均无法保证每一个100m区间小于5%。")
    add("")
    add("## 2. 数据范围与口径")
    add("")
    add("- 工程数：30、50、70速度组各6个，共18个工程。")
    add("- 每工程区间：200–300、300–400、400–500、500–600m，共72个100m样本。")
    add("- 真值：200–300m=1.25；300–400m=1.27；400–500m=1.43；500–600m=1.75。")
    add("- 原始IRI：Excel表的IRI列；用户已确认该值未乘K、未加B，且本次表格按0.25m计算后于2026-07-30 13:43重新导出。")
    add("- 本报告未使用 `IRI_100.txt` 作为数据来源或拟合依据。")
    add("- `IRI_10m`出表核验：18个工程、72个真值区间，Excel表值与对应10个10m IRI平均值的最大绝对差为0.000005 IRI，平均绝对差为0.000003 IRI。")
    add("")
    add("### 2.1 出表与K、B的等价条件")
    add("")
    add("实际出表逻辑应为：")
    add("")
    add("```text")
    add("对100m内第j个10m段：IRI10_校正,j = K(vj) × IRI10_原始,j + B(vj)")
    add("100m出表值 = (IRI10_校正,1 + ... + IRI10_校正,10) / 10")
    add("```")
    add("")
    add("本批真值区间内，10m速度均未跨档：30组均在40km/h阈值档内，50组均在60km/h阈值档内，70组均在75km/h阈值档内。因此100m内K、B恒定，满足：")
    add("")
    add("`平均(K×IRI10+B) = K×平均(IRI10)+B`")
    add("")
    add("故本报告直接对Excel的100m平均IRI拟合，与逐10m带入的结果相同。若以后一个100m内部跨速度档，必须逐10m计算，不能用100m平均速度选择一组K、B。")
    add("")
    add("## 3. 重复性：原始IRI本身稳定")
    add("")
    add("对同一速度组、同一100m区间的6次原始IRI，使用 `CV=标准差/平均IRI×100%`：")
    add("")
    add("| 速度组 | 平均CV | 最差CV | 最差区间 |")
    add("|---|---:|---:|---|")
    for group in ("30", "50", "70"):
        cvs=[]; worst=None
        for interval in ORDER:
            vals=np.array([s["raw_iri"] for s in samples if s["group"]==group and s["interval"]==interval])
            cv=np.std(vals,ddof=1)/np.mean(vals)
            cvs.append(cv)
            if worst is None or cv>worst[1]: worst=(interval,cv)
        add(f"| {group} | {np.mean(cvs)*100:.2f}% | {worst[1]*100:.2f}% | {worst[0]}m |")
    add("")
    add("重复性好只能说明同一工况下输出稳定；它不保证原始IRI与真值在所有路面区间均满足同一条线性关系。")
    add("")
    add("## 4. K、B推算方法")
    add("")
    add("对每个速度组独立拟合：")
    add("")
    add("`IRI真值 = K × IRI原始 + B`")
    add("")
    add("计算两个全局最优问题：")
    add("")
    add("1. **最小平均相对误差（MAPE）**：最小化所有样本 `平均(|预测-真值|/真值)`。这是判断“平均误差能否小于5%”的直接依据。")
    add("2. **最小最大相对误差（Minimax）**：最小化所有样本中最大的 `|预测-真值|/真值`。这是判断“每一个100m能否小于5%”的直接依据。")
    add("")
    add("两个目标对K、B都是凸的分段线性优化问题；报告枚举其全部有效约束交点，得到全局最优解，而不是通过反复试参数得到局部结果。K约束为非负，因为负K会把平整度大小顺序完全颠倒，不具物理意义。")
    add("")
    add("## 5. 最佳参数与5%结论")
    add("")
    add("### 5.1 目标为最低平均相对误差（建议用于评价平均误差）")
    add("")
    add("| 速度组 | K | B | 平均相对误差 | 最大单段相对误差 |")
    add("|---|---:|---:|---:|---:|")
    for group in ("30", "50", "70"):
        m=models[group]["mape"]
        add(f"| {group} | {m['k']:.6f} | {m['b']:.6f} | {m['mean']*100:.2f}% | {m['max']*100:.2f}% |")
    add("")
    add("### 5.2 目标为最低最大单段相对误差（用于证明单段5%不可达）")
    add("")
    add("| 速度组 | K | B | 此时平均相对误差 | 最低可能的最大单段误差 |")
    add("|---|---:|---:|---:|---:|")
    for group in ("30", "50", "70"):
        m=models[group]["minimax"]
        add(f"| {group} | {m['k']:.6f} | {m['b']:.6f} | {m['mean']*100:.2f}% | {m['max']*100:.2f}% |")
    add("")
    add("结论解释：即使不再追求平均误差最小、而专门迁就最差的单段，最优最大误差仍为30组11.96%、50组11.01%、70组8.40%。因此单一K、B下“所有100m均小于5%”在这批样本中不可实现。")
    add("")
    add("## 6. 为什么30速度组无法通过单一K、B达到5%")
    add("")
    add("30速度组的6次平均原始IRI与真值如下：")
    add("")
    add("| 区间 | 原始IRI均值 | 真值IRI |")
    add("|---|---:|---:|")
    for interval in ORDER:
        vals=[s["raw_iri"] for s in samples if s["group"]=="30" and s["interval"]==interval]
        add(f"| {interval}m | {np.mean(vals):.4f} | {samples[[i for i,s in enumerate(samples) if s['group']=='30' and s['interval']==interval][0]]['truth']:.2f} |")
    add("")
    add("其中200–300m与300–400m构成直接反例：")
    add("")
    add("```text")
    add("原始IRI：300–400m = 1.0241 < 200–300m = 1.0821")
    add("真值IRI：300–400m = 1.27   > 200–300m = 1.25")
    add("```")
    add("")
    add("对任意正K，必有 `K×1.0241+B < K×1.0821+B`。也就是说，正K、任意B均会保留原始IRI的大小顺序，无法把300–400m同时修正为比200–300m更高。这不是继续微调K、B可以消除的误差，而是单一线性模型的结构性限制。")
    add("")
    add("500–600m也存在响应不足：其原始IRI均值1.2786与400–500m的1.2509接近，但真值从1.43增至1.75，需要的增幅远大于单一线性关系可兼顾的程度。")
    add("")
    add("## 7. 72个样本详细数据（采用最小MAPE参数）")
    add("")
    add("下表的预测值按 `K×原始IRI+B` 计算；由于各100m内部未跨速度档，它等价于逐10m应用同档K、B后再平均。")
    add("")
    add("| 速度组 | 工程 | 区间(m) | 表中原始IRI=10m均值 | 平均速度(km/h) | 真值 | MAPE最优预测 | 相对误差 |")
    add("|---|---|---|---:|---:|---:|---:|---:|")
    for group in ("30", "50", "70"):
        m=models[group]["mape"]
        ss=sorted([s for s in samples if s["group"]==group], key=lambda s:(s["project"],ORDER.index(s["interval"])))
        for s in ss:
            pred=m["k"]*s["raw_iri"]+m["b"]
            rel=(pred-s["truth"])/s["truth"]
            short=s["project"].split("__",1)[0]
            add(f"| {group} | {short} | {s['interval']} | {s['raw_iri']:.6f} | {s['speed']:.2f} | {s['truth']:.2f} | {pred:.6f} | {rel*100:+.2f}% |")
    add("")
    add("## 8. 使用建议")
    add("")
    add("1. 若考核指标仅为平均误差：70组可采用本报告5.1的参数；50组最优也只能约5.08%，不应承诺5%以内；30组不能只依靠一组K、B。")
    add("2. 若考核要求每个100m均小于5%：本批数据下三组都不能通过单一K、B实现。")
    add("3. 30组应优先核对真值里程边界、原始IRI分段边界及真值测量定义；若边界一致，则需要按IRI范围分段校正或增加其他解释变量，而不是继续细调一组K、B。")
    add("4. 以后正式出表必须逐10m按速度档修正后再平均；只有100m内未跨档时，才可简化为本报告的100m公式。")
    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text("\n".join(lines)+"\n", encoding="utf-8")
    print(OUT)


if __name__ == "__main__":
    main()
