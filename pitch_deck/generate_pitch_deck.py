"""ネガティブ・チャーン & LTV向上の説得ロジック — ピッチ資料生成スクリプト

python-pptx で 16:9 / 全7スライドの .pptx を生成する。
トヨマネスタイル：
  - 色は3色のみ（濃紺・青・グレー）＋白背景
  - 各スライドのタイトルは「結論（メッセージ）」を1文で書く
  - 影・グラデーション・装飾は使わず、表と枠で情報を整理する

    pip install python-pptx
    python generate_pitch_deck.py
"""
from pathlib import Path

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

BASE_DIR = Path(__file__).resolve().parent
OUTPUT = BASE_DIR / "pitch_deck_negative_churn.pptx"
LOGO = BASE_DIR / "assets" / "taxel_logo_light.png"

# ---- 3色のみ（背景・白抜き文字の白は除く） ----
NAVY = RGBColor(0x0F, 0x17, 0x2A)   # 本文・主要要素
BLUE = RGBColor(0x25, 0x63, 0xEB)   # 強調（結論・キーワードのみ）
GRAY = RGBColor(0xE2, 0xE8, 0xF0)   # 枠・面・区切り
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

FONT = "游ゴシック"
SLIDE_W, SLIDE_H = 13.333, 7.5
MARGIN = 0.6
CONTENT_W = SLIDE_W - 2 * MARGIN
BODY_TOP = 1.95


# ---------------------------------------------------------------- helpers
def set_font(run, size, color=NAVY, bold=False):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = FONT
    rpr = run._r.get_or_add_rPr()
    for tag in ("a:ea", "a:cs"):
        el = rpr.find(qn(tag))
        if el is None:
            el = rpr.makeelement(qn(tag), {})
            rpr.append(el)
        el.set("typeface", FONT)


def fill_text(tf, lines, size, color, bold, align, anchor, spacing=1.2):
    """lines: str / [str | (text, size, color, bold)]。1要素＝1段落"""
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    if isinstance(lines, str):
        lines = [lines]
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = spacing
        t, s, c, b = line if isinstance(line, tuple) else (line, size, color, bold)
        run = p.add_run()
        run.text = t
        set_font(run, s, c, b)


def text(slide, x, y, w, h, lines, size=16, color=NAVY, bold=False,
         align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    fill_text(tb.text_frame, lines, size, color, bold, align, anchor)
    return tb


def _flat(shape):
    """テーマ由来の影・効果を持たないフラットな図形にする"""
    style = shape._element.find(qn("p:style"))
    if style is not None:
        shape._element.remove(style)


def box(slide, x, y, w, h, fill=None, line=None, lines=None, size=14, color=NAVY,
        bold=False, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE, pad=0.2,
        shape=MSO_SHAPE.RECTANGLE):
    sp = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    _flat(sp)
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid()
        sp.fill.fore_color.rgb = fill
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line
        sp.line.width = Pt(1.25)
    if lines is not None:
        fill_text(sp.text_frame, lines, size, color, bold, align, anchor)
        tf = sp.text_frame
        tf.margin_left = tf.margin_right = Inches(pad)
    return sp


def chevron(slide, x, y, w=0.28, h=0.4, fill=NAVY):
    """ステップ間の矢印（右向き三角）"""
    sp = slide.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, Inches(x), Inches(y),
                                Inches(h), Inches(w))
    sp.rotation = 90
    _flat(sp)
    sp.fill.solid()
    sp.fill.fore_color.rgb = fill
    sp.line.fill.background()
    return sp


def hline(slide, x1, y, x2, color=GRAY, weight=1.25):
    ln = slide.shapes.add_connector(1, Inches(x1), Inches(y), Inches(x2), Inches(y))
    _flat(ln)
    ln.line.color.rgb = color
    ln.line.width = Pt(weight)
    return ln


def header(slide, n, label, message, size=24):
    """左上：資料上の位置づけ（ラベル）／タイトル：結論メッセージ"""
    text(slide, MARGIN, 0.45, 10.4, 0.3, label, size=12, color=BLUE, bold=True)
    text(slide, MARGIN, 0.8, 10.8, 0.9, message, size=size, bold=True)
    if LOGO.exists():
        slide.shapes.add_picture(str(LOGO), Inches(SLIDE_W - MARGIN - 1.3), Inches(0.45),
                                 height=Inches(0.36))
    hline(slide, MARGIN, SLIDE_H - 0.55, SLIDE_W - MARGIN)
    text(slide, MARGIN, SLIDE_H - 0.45, 6, 0.25, "taXel｜投資家向け補足資料", size=9)
    text(slide, SLIDE_W - MARGIN - 0.5, SLIDE_H - 0.45, 0.5, 0.25, str(n), size=9,
         align=PP_ALIGN.RIGHT)


def point(slide, y, message, h=0.7):
    """スライド下部の結論ボックス（濃紺・白抜き）"""
    box(slide, MARGIN, y, 1.3, h, fill=BLUE, lines="結論", size=14, color=WHITE, bold=True,
        align=PP_ALIGN.CENTER)
    box(slide, MARGIN + 1.3, y, CONTENT_W - 1.3, h, fill=NAVY, lines=message, size=17,
        color=WHITE, bold=True, pad=0.3)


def table(slide, x, y, col_w, row_h, rows, size=14, highlight_col=None):
    """rows[0] はヘッダー。セルは枠線グレーのフラットな矩形で描く"""
    for r, row in enumerate(rows):
        cx = x
        for c, cell in enumerate(row):
            if r == 0:
                fill, color, bold = (BLUE if c == highlight_col else NAVY), WHITE, True
            elif c == 0:
                fill, color, bold = GRAY, NAVY, True
            elif c == highlight_col:
                fill, color, bold = WHITE, BLUE, True
            else:
                fill, color, bold = WHITE, NAVY, False
            box(slide, cx, y + r * row_h, col_w[c], row_h, fill=fill, line=GRAY, lines=cell,
                size=size, color=color, bold=bold)
            cx += col_w[c]


def new_slide(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])  # Blank
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = WHITE
    return s


# ---------------------------------------------------------------- slides
def slide1_title(prs):
    s = new_slide(prs)
    if LOGO.exists():
        s.shapes.add_picture(str(LOGO), Inches(MARGIN), Inches(0.6), height=Inches(0.55))
    text(s, MARGIN, 2.0, 11, 0.35, "投資家向け補足資料｜ユニットエコノミクス", size=14,
         color=BLUE, bold=True)
    text(s, MARGIN, 2.45, CONTENT_W, 1.0, "ネガティブ・チャーン & LTV向上の説得ロジック",
         size=36, bold=True)
    text(s, MARGIN, 3.45, CONTENT_W, 0.5,
         "税理士チャネルが生む「構造的な」収益拡大を4つの軸で証明", size=18)

    axes = ["構造的クロスセル", "スイッチングコスト", "LTV / CAC 拡大", "シナリオ分析"]
    gap = 0.2
    w = (CONTENT_W - 3 * gap) / 4
    for i, label in enumerate(axes):
        x = MARGIN + i * (w + gap)
        box(s, x, 4.9, 0.6, 0.8, fill=NAVY, lines=f"{i + 1}", size=20, color=WHITE,
            bold=True, align=PP_ALIGN.CENTER)
        box(s, x + 0.6, 4.9, w - 0.6, 0.8, fill=GRAY, lines=label, size=14, bold=True,
            pad=0.15)
    hline(s, MARGIN, SLIDE_H - 0.55, SLIDE_W - MARGIN)
    text(s, MARGIN, SLIDE_H - 0.45, 8, 0.25, "taXel（タクセル）｜2026年9月", size=9)


def slide2_overview(prs):
    s = new_slide(prs)
    header(s, 2, "概要｜投資家が求める「Why & How」の証明",
           "ネガティブ・チャーンは「期待」ではなく「構造」で説明できる")

    box(s, MARGIN, BODY_TOP, CONTENT_W, 0.65, fill=GRAY, pad=0.3,
        lines=[("投資家の問い：なぜ解約が起きても、既存顧客からの売上は伸び続けるのか？",
                17, NAVY, True)])
    rows = [
        ["軸", "問い", "答え（要点）"],
        ["1  構造的クロスセル", "How：どう伸びるか", "税理士経由で利用プロダクト数が拡大"],
        ["2  スイッチングコスト", "Why：なぜ辞めないか", "業務インフラ化 ＋ 税理士によるCS"],
        ["3  LTV / CAC", "Proof：数字で示せるか", "初年度3倍 → 3年目6倍以上"],
        ["4  シナリオ分析", "Risk：下振れ時は", "保守的な下限とアップサイドを両提示"],
    ]
    table(s, MARGIN, BODY_TOP + 0.95, [3.4, 3.4, CONTENT_W - 6.8], 0.6, rows, size=15,
          highlight_col=2)
    point(s, 5.95, "4軸すべてが「税理士チャネル」という1つの構造から生まれる")


def slide3_crosssell(prs):
    s = new_slide(prs)
    header(s, 3, "軸1｜構造的なクロスセルの仕組み（税理士チャネルの横展開）",
           "1事務所の獲得が、同一税理士経由のクロスセルへ連鎖する")

    steps = [("STEP 1", "税理士事務所を獲得"), ("STEP 2", "1プロダクトで導入"),
             ("STEP 3", "信頼関係の構築"), ("STEP 4", "他プロダクトへ展開")]
    gap = 0.5
    w = (CONTENT_W - 3 * gap) / 4
    for i, (st, t) in enumerate(steps):
        x = MARGIN + i * (w + gap)
        last = i == len(steps) - 1
        box(s, x, BODY_TOP, w, 1.0, fill=BLUE if last else GRAY, pad=0.25,
            lines=[(st, 11, WHITE if last else BLUE, True),
                   (t, 16, WHITE if last else NAVY, True)])
        if not last:
            chevron(s, x + w + 0.1, BODY_TOP + 0.36)

    # 左：プロダクト数の積み上げ（イメージ）
    py, lw, ph = 3.3, 6.2, 2.35
    box(s, MARGIN, py, lw, ph, line=GRAY)
    text(s, MARGIN + 0.25, py + 0.15, lw - 0.5, 0.3, "1税理士あたり利用プロダクト数（イメージ）",
         size=12, bold=True)
    names = ["プロダクトA", "＋プロダクトB", "＋プロダクトC"]
    fills = [NAVY, BLUE, GRAY]
    colors = [WHITE, WHITE, NAVY]
    cw, bh, base = 1.5, 0.4, py + ph - 0.45
    for i, yr in enumerate(["1年目", "2年目", "3年目"]):
        cx = MARGIN + 0.45 + i * (cw + 0.45)
        for j in range(i + 1):
            box(s, cx, base - (j + 1) * (bh + 0.04), cw, bh, fill=fills[j], lines=names[j],
                size=11, color=colors[j], bold=True, align=PP_ALIGN.CENTER, pad=0.05)
        text(s, cx, base + 0.08, cw, 0.3, yr, size=11, align=PP_ALIGN.CENTER)

    # 右：因果の整理
    rx = MARGIN + lw + 0.3
    rw = SLIDE_W - MARGIN - rx
    rows = [["事象", "影響"],
            ["単一プロダクトの解約", "発生しても影響は限定的"],
            ["1税理士あたりARPU", "年々拡大"],
            ["利用プロダクト数", "年々拡大"]]
    table(s, rx, py, [2.5, rw - 2.5], ph / 4, rows, size=14, highlight_col=1)
    point(s, 5.95, "拡大が解約を上回り、全体としてネガティブ・チャーンを達成")


def slide4_switching(prs):
    s = new_slide(prs)
    header(s, 4, "軸2｜スイッチングコストと高定着性の根拠",
           "導入6か月で業務インフラ化し、税理士CSが定着を支える")

    colw = (CONTENT_W - 0.3) / 2
    lx, rx = MARGIN, MARGIN + colw + 0.3
    for x, tag, title in [(lx, "A", "業務インフラ化"), (rx, "B", "税理士によるサクセスサポート")]:
        box(s, x, BODY_TOP, 0.6, 0.6, fill=NAVY, lines=tag, size=18, color=WHITE, bold=True,
            align=PP_ALIGN.CENTER)
        box(s, x + 0.6, BODY_TOP, colw - 0.6, 0.6, fill=GRAY, lines=title, size=17, bold=True)

    # A：タイムライン
    ty = BODY_TOP + 0.85
    x0, x1 = lx + 0.5, lx + colw - 0.5
    hline(s, x0, ty + 0.5, x1, color=NAVY, weight=2)
    for label, frac, emph in [("導入", 0.0, False), ("6か月", 0.5, True), ("定着", 1.0, False)]:
        cx = x0 + frac * (x1 - x0)
        box(s, cx - 0.12, ty + 0.38, 0.24, 0.24, fill=BLUE if emph else NAVY,
            shape=MSO_SHAPE.OVAL)
        text(s, cx - 0.8, ty, 1.6, 0.3, label, size=13, color=BLUE if emph else NAVY,
             bold=True, align=PP_ALIGN.CENTER)
    text(s, lx, ty + 0.85, colw, 0.3, "日常の業務フローに完全に組み込み", size=14,
         align=PP_ALIGN.CENTER)
    cw = (colw - 0.4) / 3
    for i, c in enumerate(["経理", "決済", "リスク管理"]):
        box(s, lx + i * (cw + 0.2), ty + 1.3, cw, 0.5, line=NAVY, lines=c, size=14,
            bold=True, align=PP_ALIGN.CENTER)
    box(s, lx, ty + 2.05, colw, 0.55, line=BLUE,
        lines="スイッチングコスト上昇 → 解約率が急減", size=15, color=BLUE, bold=True,
        align=PP_ALIGN.CENTER)

    # B：表
    rows = [["税理士の役割", "効果"],
            ["決算・財務を把握", "顧客状況を最も理解"],
            ["継続的に関与・提案", "顧問業務の中で活用を促進"],
            ["実質的なCS担当", "CSコスト抑制 × 高定着"]]
    table(s, rx, BODY_TOP + 0.85, [2.6, colw - 2.6], 0.65, rows, size=14, highlight_col=1)
    point(s, 5.95, "定着の仕組みがプロダクトと税理士の両面に組み込まれている")


def slide5_unit_economics(prs):
    s = new_slide(prs)
    header(s, 5, "軸3｜ユニットエコノミクス（LTV / CAC）の拡大ストーリー",
           "LTV / CAC は初年度3倍から、3年目に6倍以上へ拡大する")

    # 左：棒グラフ
    cw, ch = 5.2, 3.7
    box(s, MARGIN, BODY_TOP, cw, ch, line=GRAY)
    text(s, MARGIN + 0.25, BODY_TOP + 0.15, 2.5, 0.3, "LTV / CAC 倍率", size=13, bold=True)
    text(s, MARGIN + cw - 2.35, BODY_TOP + 0.15, 2.1, 0.3, "3年目は6倍以上", size=12,
         color=BLUE, bold=True, align=PP_ALIGN.RIGHT)
    data = CategoryChartData()
    data.categories = ["初年度", "3年目"]
    data.add_series("LTV/CAC", (3, 6))
    chart = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(MARGIN + 0.25),
                               Inches(BODY_TOP + 0.5), Inches(cw - 0.5), Inches(ch - 0.65),
                               data).chart
    chart.has_legend = False
    chart.has_title = False
    plot = chart.plots[0]
    plot.gap_width = 100
    plot.has_data_labels = True
    dl = plot.data_labels
    dl.number_format = '0"倍"'
    dl.number_format_is_linked = False
    dl.position = XL_LABEL_POSITION.OUTSIDE_END
    dl.font.size = Pt(20)
    dl.font.bold = True
    dl.font.color.rgb = NAVY
    dl.font.name = FONT
    ser = plot.series[0]
    ser.format.fill.solid()
    ser.format.fill.fore_color.rgb = BLUE
    p0 = ser.points[0]
    p0.format.fill.solid()
    p0.format.fill.fore_color.rgb = NAVY
    va = chart.value_axis
    va.visible = False
    va.has_major_gridlines = False
    va.minimum_scale, va.maximum_scale = 0, 7.5
    ca = chart.category_axis
    ca.tick_labels.font.size = Pt(13)
    ca.tick_labels.font.color.rgb = NAVY
    ca.tick_labels.font.name = FONT
    ca.format.line.color.rgb = GRAY

    # 右：ドライバー表 ＋ Payback
    rx = MARGIN + cw + 0.3
    rw = SLIDE_W - MARGIN - rx
    rows = [["ドライバー", "打ち手", "効果"],
            ["CAC", "税理士チャネルの活用", "劇的に低く抑制"],
            ["初期解約率", "導入初期の定着", "早期離脱を抑制"],
            ["クロスセル", "同一税理士経由", "1社あたりLTV拡大"]]
    table(s, rx, BODY_TOP, [1.8, 2.4, rw - 4.2], 0.55, rows, size=14, highlight_col=2)
    py = BODY_TOP + 2.45
    box(s, rx, py, 1.8, 1.25, fill=NAVY, lines=["Payback", "Period"], size=15, color=WHITE,
        bold=True, align=PP_ALIGN.CENTER)
    box(s, rx + 1.8, py, rw - 1.8, 1.25, line=GRAY, pad=0.3,
        lines=[("CAC回収期間が極めて短い", 17, NAVY, True),
               ("回収後の利益がストックとして積み上がる", 14, BLUE, True)])
    point(s, 5.95, "低CAC × 解約率低下 × クロスセルで、倍率は年々改善する")


def slide6_scenarios(prs):
    s = new_slide(prs)
    header(s, 6, "軸4｜シナリオ分析による実現可能性（ベースライン vs アップサイド）",
           "保守的な下限に対し、実効LTVは＋30〜50%の上振れ余地")

    tw = 7.6
    rows = [["項目", "ベースライン", "アップサイド"],
            ["位置づけ", "保守的な下限値", "ネガティブ・チャーン事例"],
            ["解約率", "2.0% 固定", "初期解約率が低下"],
            ["クロスセル", "なし", "税理士経由で進展"],
            ["実効LTV", "100（基準）", "130〜150（+30〜50%）"]]
    table(s, MARGIN, BODY_TOP, [1.8, 2.2, tw - 4.0], 0.74, rows, size=15, highlight_col=2)

    # 右：実効LTVの比較バー（シェイプ）
    vx = MARGIN + tw + 0.3
    vw = SLIDE_W - MARGIN - vx
    vh = 3.7
    box(s, vx, BODY_TOP, vw, vh, line=GRAY)
    text(s, vx + 0.25, BODY_TOP + 0.15, vw - 0.5, 0.3, "実効LTV（ベースライン＝100）", size=13,
         bold=True)
    base = BODY_TOP + vh - 0.5
    unit = 2.3 / 150
    bw = 1.1
    b1, b2 = vx + 0.55, vx + vw - 0.55 - bw
    box(s, b1, base - 100 * unit, bw, 100 * unit, fill=NAVY)
    box(s, b2, base - 130 * unit, bw, 130 * unit, fill=BLUE)
    box(s, b2, base - 150 * unit, bw, 20 * unit, fill=GRAY, line=BLUE)
    text(s, b1 - 0.3, base - 100 * unit - 0.42, bw + 0.6, 0.35, "100", size=18, bold=True,
         align=PP_ALIGN.CENTER)
    text(s, b2 - 0.4, base - 150 * unit - 0.42, bw + 0.8, 0.35, "130〜150", size=18,
         color=BLUE, bold=True, align=PP_ALIGN.CENTER)
    hline(s, vx + 0.3, base, vx + vw - 0.3, color=NAVY)
    for bx, lab in [(b1, "ベースライン"), (b2, "アップサイド")]:
        text(s, bx - 0.3, base + 0.08, bw + 0.6, 0.3, lab, size=12, align=PP_ALIGN.CENTER)
    point(s, 5.95, "下限で計画を語り、上振れは構造的な余地として示す")


def slide7_conclusion(prs):
    s = new_slide(prs)
    header(s, 7, "結論",
           "単体SaaSではなく「税理士プラットフォーム」としての構造的優位性", size=22)

    cols = [("獲得", "低CAC", "税理士チャネルで効率的に獲得"),
            ("定着", "高い定着性", "業務インフラ化＋税理士CS"),
            ("拡張", "クロスセル", "同一税理士経由で横展開"),
            ("結果", "ネガティブ・チャーン", "LTVが時間とともに拡大")]
    gap = 0.5
    w = (CONTENT_W - 3 * gap) / 4
    for i, (kw, title, body) in enumerate(cols):
        x = MARGIN + i * (w + gap)
        last = i == len(cols) - 1
        box(s, x, BODY_TOP, w, 0.5, fill=BLUE if last else NAVY, lines=kw, size=15,
            color=WHITE, bold=True, align=PP_ALIGN.CENTER)
        box(s, x, BODY_TOP + 0.5, w, 1.5, fill=None if last else GRAY,
            line=BLUE if last else None, pad=0.15,
            lines=[(title, 16, BLUE if last else NAVY, True), (body, 13, NAVY, False)])
        if not last:
            chevron(s, x + w + 0.11, BODY_TOP + 1.0)

    rows = [["比較軸", "単体SaaS", "税理士プラットフォーム"],
            ["顧客獲得", "1社ずつ個別に獲得", "税理士1事務所から顧問先へ波及"],
            ["収益の伸び", "1プロダクトに依存", "複数プロダクトで拡大"]]
    table(s, MARGIN, BODY_TOP + 2.3, [2.4, 3.8, CONTENT_W - 6.2], 0.5, rows, size=14,
          highlight_col=2)
    point(s, 5.95, "税理士1事務所の獲得が、顧問先 × 複数プロダクトの収益へ波及")


def main():
    prs = Presentation()
    prs.slide_width = Inches(SLIDE_W)
    prs.slide_height = Inches(SLIDE_H)
    for build in (slide1_title, slide2_overview, slide3_crosssell, slide4_switching,
                  slide5_unit_economics, slide6_scenarios, slide7_conclusion):
        build(prs)
    prs.save(OUTPUT)
    print(f"Saved: {OUTPUT}")


if __name__ == "__main__":
    main()
