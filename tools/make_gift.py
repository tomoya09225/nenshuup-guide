"""公式LINEの登録特典（PDF 3点）を作る。

    python tools/make_gift.py            # gift/ に HTML と PDF を書き出す

① 年収1000万円超 企業リスト：nenshu-up-automation/content/edinet/salary.csv（各社 有価証券報告書「従業員の状況」）
② 年収の相場 早見表：国税庁 民間給与実態統計調査（数字は nenshu-up-automation の qc.json で本文と照合済み）
③ 年収を上げる転職 チェックリスト：一般的な転職の進め方（数字・事例は使わない）
PDF は Chrome のヘッドレス印刷で作る。
"""
import csv
import html
import re
import subprocess
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'gift'
EDINET = ROOT.parent / 'nenshu-up-automation' / 'content' / 'edinet' / 'salary.csv'
CHROME = r'C:\Program Files\Google\Chrome\Application\chrome.exe'
LINE_ID = '@616nxycf'
CONSULT = '年収UPキャリア相談'
# EDINET の取り込みで単位がずれた値（数億円など）が数件あるため、キーエンス（2,178万円）を大きく超える値は除く
MAX_YEN = 30_000_000
CREDIT = '人材業界8年目・キャリア相談1,000名以上の転職のプロが作成'

CSS = """
@page{size:A4;margin:14mm 13mm 14mm}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:"Noto Sans JP",sans-serif;color:#1b1f27;font-size:10.5pt;line-height:1.6;-webkit-print-color-adjust:exact;print-color-adjust:exact}
.serif{font-family:"Shippori Mincho B1",serif}
.cover{height:265mm;background:#0d121c;color:#f3eee3;border-radius:6mm;padding:22mm 16mm;position:relative;overflow:hidden;page-break-after:always;
  background-image:radial-gradient(120% 60% at 50% -10%,rgba(201,164,92,.22),transparent 60%),repeating-linear-gradient(135deg,rgba(201,164,92,.05) 0 1px,transparent 1px 12px)}
.cover:before{content:"";position:absolute;inset:5mm;border:0.4mm solid rgba(201,164,92,.45);border-radius:4mm}
.cover .tag{display:inline-block;border:0.3mm solid #c9a45c;color:#ead39b;font-size:10pt;letter-spacing:.2em;padding:1.5mm 4mm;border-radius:10mm}
.cover h1{font-family:"Shippori Mincho B1",serif;font-size:34pt;line-height:1.35;margin:14mm 0 8mm;font-weight:800}
.cover h1 b{color:#e3c27a}
.cover p{font-size:12pt;color:#c9cdd6;line-height:1.9}
.cover .foot{position:absolute;left:16mm;right:16mm;bottom:20mm;font-size:10pt;color:#a3abba;border-top:0.3mm solid rgba(201,164,92,.4);padding-top:5mm}
.cover .foot b{display:block;font-family:"Shippori Mincho B1",serif;font-size:16pt;color:#f3eee3;margin-bottom:1mm}
h2{font-family:"Shippori Mincho B1",serif;font-size:17pt;font-weight:800;margin:0 0 3mm;padding-left:3.5mm;border-left:1.6mm solid #c9a45c}
h3{font-size:12pt;font-weight:900;margin:5mm 0 2mm;color:#0d121c}
.lead{color:#545b68;font-size:9.5pt;margin-bottom:4mm}
.page{page-break-after:always}
.page:last-child{page-break-after:auto}
table{width:100%;border-collapse:collapse;font-size:9.3pt;margin-bottom:3mm}
th{background:#0d121c;color:#ead39b;font-weight:700;padding:1.6mm 2mm;text-align:left}
td{padding:1.4mm 2mm;border-bottom:0.25mm solid #e3dfd5}
tr:nth-child(even) td{background:#faf8f3}
td.num,th.num{text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap}
td.big{font-weight:900;color:#8d6a2c}
.ind{break-inside:avoid;margin-bottom:3mm}
.ind h3{display:flex;justify-content:space-between;align-items:baseline;border-bottom:0.5mm solid #c9a45c;padding-bottom:1mm}
.ind h3 span{font-size:9pt;color:#7c8496;font-weight:500}
.note{font-size:8.5pt;color:#6b7280;line-height:1.7;margin-top:3mm}
.box{border:0.4mm solid #c9a45c;background:#fbf7ee;border-radius:3mm;padding:4mm 5mm;margin:4mm 0;break-inside:avoid}
.box b{color:#8d6a2c}
.cta{background:#0d121c;color:#f3eee3;border-radius:4mm;padding:8mm 8mm;margin-top:6mm;break-inside:avoid}
.cta h3{color:#ead39b;font-family:"Shippori Mincho B1",serif;font-size:15pt;margin:0 0 2mm}
.cta p{color:#c9cdd6;font-size:10pt}
.cta .kw{display:inline-block;background:#06c755;color:#fff;font-weight:900;font-size:13pt;padding:2mm 6mm;border-radius:2mm;margin:3mm 0 1mm}
.check{list-style:none;margin:0 0 2mm}
.check li{padding:1.6mm 0 1.6mm 8mm;border-bottom:0.25mm dashed #d9d3c4;position:relative;break-inside:avoid}
.check li:before{content:"";position:absolute;left:0;top:2.3mm;width:4mm;height:4mm;border:0.4mm solid #8d6a2c;border-radius:0.8mm}
.check li b{display:block}
.check li span{color:#545b68;font-size:9.3pt}
.grid2{display:grid;grid-template-columns:1fr 1fr;gap:5mm}
.bar{height:3mm;background:linear-gradient(90deg,#a9803a,#ead39b);border-radius:2mm}
"""

HEAD = ('<!doctype html><html lang="ja"><head><meta charset="utf-8"><title>{t}</title>'
        '<link href="https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@400;500;700;900&family=Shippori+Mincho+B1:wght@700;800&display=swap" rel="stylesheet">'
        '<style>' + CSS + '</style></head><body>')


def cover(tag, title, sub):
    return (f'<div class="cover"><span class="tag">{tag}</span><h1>{title}</h1><p>{sub}</p>'
            f'<div class="foot"><b>年収UP研究所</b>{CREDIT}<br>Instagram @nenshu_up__lab</div></div>')


def cta():
    return (f'<div class="cta"><h3>あなたの経歴なら、どの会社を狙える？</h3>'
            f'<p>公式LINEで、このキーワードを送ってください。年収UPのプロが、あなたの経歴をもとに直接相談にのります（無料・30分・LINEかオンライン面談）。</p>'
            f'<div class="kw">{CONSULT}</div><p style="font-size:8.5pt;color:#a3abba">転職を決めていなくても大丈夫です。公式LINE {LINE_ID}</p></div>')


def clean(name):
    return re.sub(r'\s+', '', name.replace('株式会社', '').replace('（株）', '')).strip()


def man(yen):
    return f'{int(yen) / 10000:,.1f}'


# ---------- ① 年収1000万円超 企業リスト ----------
def gift_companies():
    rows = [r for r in csv.DictReader(EDINET.open(encoding='utf-8'))
            if r['avg_salary_yen'] and 10_000_000 <= int(r['avg_salary_yen']) <= MAX_YEN and int(r['employees'] or 0) >= 300]
    rows.sort(key=lambda r: -int(r['avg_salary_yen']))
    by_ind = defaultdict(list)
    for r in rows:
        by_ind[r['industry']].append(r)
    order = sorted(by_ind, key=lambda k: (-len(by_ind[k]), -int(by_ind[k][0]['avg_salary_yen'])))
    periods = sorted({r['periodEnd'][:7] for r in rows})
    top = ''.join(
        f'<tr><td class="num">{i}</td><td><b>{html.escape(clean(r["filerName"]))}</b></td><td>{html.escape(r["industry"])}</td>'
        f'<td class="num big">{man(r["avg_salary_yen"])}</td><td class="num">{float(r["avg_age"]):.1f}歳</td>'
        f'<td class="num">{int(r["employees"]):,}人</td></tr>' for i, r in enumerate(rows[:30], 1))
    inds = ''.join(
        f'<div class="ind"><h3>{html.escape(k)}<span>{len(by_ind[k])}社</span></h3><table>'
        '<tr><th>会社名</th><th class="num">平均年収（万円）</th><th class="num">平均年齢</th><th class="num">平均勤続</th><th class="num">従業員数</th></tr>'
        + ''.join(f'<tr><td>{html.escape(clean(r["filerName"]))}</td><td class="num big">{man(r["avg_salary_yen"])}</td>'
                  f'<td class="num">{float(r["avg_age"]):.1f}歳</td><td class="num">{float(r["avg_tenure_years"] or 0):.1f}年</td>'
                  f'<td class="num">{int(r["employees"]):,}人</td></tr>' for r in by_ind[k]) + '</table></div>'
        for k in order)
    body = (HEAD.replace('{t}', '年収1000万円超 企業リスト')
            + cover('年収UP研究所 特典 ①', f'平均年収<b>1,000万円超</b>の<br>企業リスト {len(rows)}社',
                    '業界別・有価証券報告書の実数。<br>「年収1000万円」が現実にある会社を、<br>業界ごとに一覧にしました。')
            + '<div class="page"><h2>このリストの見方</h2>'
            '<div class="box"><b>数字の出どころ</b><br>各社の有価証券報告書「従業員の状況」に載っている、提出会社（その会社単体）の平均年間給与・平均年齢・平均勤続年数・従業員数です。'
            f'金融庁の EDINET から取得しています（決算期 {periods[0]}〜{periods[-1]}）。</div>'
            '<div class="box"><b>読むときの注意</b><br>・全社員の平均です。職種・年齢・等級で大きく変わります。<br>'
            '・平均年齢が高い会社ほど平均年収も高く出やすいので、<b>平均年齢と一緒に</b>見てください。<br>'
            '・持株会社など社員の少ない会社は高く出やすいため、<b>従業員300人以上</b>の会社だけを載せています。<br>'
            '・賞与・残業代を含むかどうかは各社の注記によります。</div>'
            '<div class="box"><b>使い方</b><br>①自分の業界の会社を探す　②隣の業界で、自分の経験が活きそうな会社を探す　③気になる会社の採用ページで、自分の職種の求人があるかを見る。<br>'
            'どこを狙えるかは経歴しだいです。迷ったら公式LINEで相談してください。</div>'
            '<h2 style="margin-top:7mm">平均年収 TOP30</h2><table><tr><th class="num">順位</th><th>会社名</th><th>業種</th><th class="num">平均年収（万円）</th><th class="num">平均年齢</th><th class="num">従業員数</th></tr>'
            + top + '</table></div>'
            + f'<div class="page"><h2>業界別の一覧（{len(order)}業界・{len(rows)}社）</h2><p class="lead">業種は東京証券取引所の33業種区分。各業界の中は平均年収の高い順。</p>' + inds
            + cta() + '<p class="note">出典：各社 有価証券報告書「第一部 企業情報 第1 企業の概況 5 従業員の状況」（提出会社）。EDINET（金融庁）より取得。'
            '掲載企業との提携や、求人の紹介・内定を約束するものではありません。</p></div></body></html>')
    return 'company-list', body, len(rows)


# ---------- ② 年収の相場 早見表 ----------
AGE = [('20〜24歳', 277), ('25〜29歳', 407), ('30〜34歳', 449), ('35〜39歳', 482), ('40〜44歳', 516), ('45〜49歳', 540), ('50〜54歳', 559), ('55〜59歳', 572)]
IND = [('電気・ガス・熱供給・水道業', 854.4), ('金融業・保険業', 725.6), ('情報通信業', 681.9), ('建設業', 590.3), ('製造業', 573.0),
       ('学術研究・専門技術サービス・教育', 567.9), ('不動産業・物品賃貸業', 512.5), ('複合サービス事業', 509.5), ('運輸業・郵便業', 482.5),
       ('医療・福祉', 424.0), ('卸売業・小売業', 413.4), ('サービス業', 401.7), ('農林水産・鉱業', 393.7), ('宿泊業・飲食サービス業', 275.8)]


def gift_market():
    amax = max(v for _, v in AGE)
    imax = max(v for _, v in IND)
    age = ''.join(f'<tr><td><b>{a}</b></td><td class="num big">{v}万円</td><td style="width:55%"><div class="bar" style="width:{v / amax * 100:.0f}%"></div></td></tr>' for a, v in AGE)
    ind = ''.join(f'<tr><td>{a}</td><td class="num big">{v:,.0f}万円</td><td style="width:45%"><div class="bar" style="width:{v / imax * 100:.0f}%"></div></td></tr>' for a, v in IND)
    body = (HEAD.replace('{t}', '年収の相場 早見表')
            + cover('年収UP研究所 特典 ②', '自分の年収は<br><b>高い？低い？</b><br>相場 早見表', '年齢別・業種別の平均給与を1枚に。<br>国税庁の統計で、今の年収の位置を確かめましょう。')
            + '<div class="page"><h2>年齢別の平均給与</h2><p class="lead">1年を通じて勤務した給与所得者（正社員以外も含む・賞与込み）の平均。国税庁「令和6年分 民間給与実態統計調査」第10表（給与総額÷人数で算出）。</p>'
            '<table><tr><th>年齢</th><th class="num">平均給与</th><th></th></tr>' + age + '</table>'
            '<h2 style="margin-top:6mm">業種別の平均給与</h2><p class="lead">国税庁「令和7年分 民間給与実態統計調査結果について」業種別の平均給与（千円→万円に換算）。全体の平均は487万円。</p>'
            '<table><tr><th>業種</th><th class="num">平均給与</th><th></th></tr>' + ind + '</table>'
            '<div class="box"><b>見るときのコツ</b><br>・パート・アルバイトも含む平均なので、正社員の相場より低めに出ます。<br>'
            '・同じ業種でも、会社の規模や職種で大きく変わります。会社ごとの数字は特典①の企業リストで確かめてください。<br>'
            '・「平均より低い」と分かったら、同じ職種の求人の年収と並べてみるのが次の一歩です。</div>'
            + cta() + '<p class="note">出典：国税庁「令和6年分 民間給与実態統計調査」第10表、「令和7年分 民間給与実態統計調査結果について」（令和8年9月）p.4。</p></div></body></html>')
    return 'market-table', body, None


# ---------- ③ 年収を上げる転職 チェックリスト ----------
CHECK = [
    ('転職を考えはじめたら', [
        ('自分の市場価値（相場）を知る', '同じ年齢・職種の相場と比べる。年収UP研究所の無料診断（8問・約1分）でも確かめられます'),
        ('「なぜ年収を上げたいか」を一言にする', '生活のため／評価への不満／将来の不安など。理由で選ぶべき会社が変わります'),
        ('年収以外に譲れない条件を3つまで決める', '勤務地・働き方・仕事内容など。多すぎると選べなくなります'),
    ]),
    ('求人票を見るとき', [
        ('「想定年収」の下限を見る', '想定年収は幅で書かれることが多く、最初は下限に近い提示になることがあります'),
        ('固定残業代が含まれていないか確かめる', '「月◯時間分の固定残業代を含む」とあれば、基本給は見た目より低くなります'),
        ('賞与の「実績」を確かめる', '「年2回」だけでなく、昨年の実績（何か月分か）を面接で聞くと安心です'),
        ('平均年収は平均年齢と一緒に見る', '平均年齢が高い会社ほど平均年収も高く出やすくなります'),
    ]),
    ('職務経歴書・面接', [
        ('成果を数字で書く', '売上・達成率・件数・削減した時間など。「頑張った」より数字が伝わります'),
        ('担当した役割と規模を書く', 'チームの人数・取引先の規模・予算など、任された大きさが分かるように'),
        ('転職理由は「前向きな目的」で話す', '不満だけでなく「次に何をしたいか」につなげる'),
        ('逆質問は「役割・評価・成長」の3つから', '入社後に期待される役割／評価される成果／活躍している人の共通点'),
    ]),
    ('年収交渉', [
        ('交渉は内定後・条件提示のタイミングで', '選考の途中で年収の話ばかりすると、印象を落とすことがあります'),
        ('希望額には根拠をそえる', '今の年収（賞与込み）・同じ職種の相場・他社の選考状況など'),
        ('「総額」で比べる', '基本給・賞与・各種手当・残業代・退職金・福利厚生まで含めて比べる'),
        ('一人で抱えない', '転職エージェントを通すと、条件の交渉を代わりに進めてもらえることもあります'),
    ]),
    ('内定をもらったら', [
        ('条件は書面（労働条件通知書）で確かめる', '口頭の説明と違いがないか、年収の内訳まで確認する'),
        ('昇給と評価の仕組みを聞く', '入社時の年収だけでなく、3年後にどうなるかが大切です'),
        ('今の会社の引き止めに備える', '辞める理由と入社日を先に決めておくと、ぶれずに進められます'),
    ]),
]


def gift_checklist():
    secs = ''.join(f'<h3>{html.escape(t)}</h3><ul class="check">' + ''.join(
        f'<li><b>{html.escape(a)}</b><span>{html.escape(b)}</span></li>' for a, b in items) + '</ul>' for t, items in CHECK)
    n = sum(len(i) for _, i in CHECK)
    body = (HEAD.replace('{t}', '年収を上げる転職 チェックリスト')
            + cover('年収UP研究所 特典 ③', f'年収を上げる転職<br><b>チェックリスト{n}</b>', '転職を考えはじめてから内定まで。<br>年収で損をしないための確認ポイントを、<br>順番どおりにまとめました。')
            + f'<div class="page"><h2>チェックリスト {n}</h2><p class="lead">できたものに印をつけながら進めてください。</p>' + secs
            + cta() + '<p class="note">一般的な転職の進め方をまとめたものです。個別の状況によって最適な進め方は変わります。</p></div></body></html>')
    return 'checklist', body, n


def main():
    OUT.mkdir(exist_ok=True)
    for make in (gift_companies, gift_market, gift_checklist):
        name, body, n = make()
        src = OUT / f'{name}.html'
        src.write_text(body, encoding='utf-8')
        pdf = OUT / f'{name}.pdf'
        subprocess.run([CHROME, '--headless=new', '--disable-gpu', '--no-pdf-header-footer', '--virtual-time-budget=8000',
                        f'--print-to-pdf={pdf}', src.as_uri()], check=True, capture_output=True)
        print(name, n if n is not None else '', pdf.stat().st_size // 1024, 'KB')


if __name__ == '__main__':
    main()
