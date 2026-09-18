import pandas as pd
import streamlit as st
import requests
from datetime import datetime, timedelta

st.set_page_config(page_title="券商分點成本追蹤", page_icon="🏦", layout="wide")
st.title("📊 籌碼與期貨市場追蹤儀表板")
tab_broker, tab_foreign, tab_futures, tab_pelosi, tab_mops = st.tabs(["🏦 券商分點", "🌍 外資券商", "📈 期貨市場", "🏛️ Pelosi 持股", "📢 台股重大訊息"])

with tab_broker:
    st.header("台股券商分點成本追蹤")    
st.caption("用每日分點買進／賣出張數與均價，推估券商分點的剩餘庫存與平均成本。")

st.warning("公開分點資料代表該券商分點彙總交易，不等於單一主力或單一投資人的真實持倉。成本為研究用推估值。")

st.subheader("🔥 券商買賣前 10 名股票")
rank_file = st.file_uploader("上傳券商買賣排行 CSV（可選）", type=["csv"], key="rank_file")
if rank_file:
    rank = pd.read_csv(rank_file)
    buy_col = next((x for x in ["買超張數","買超","淨買超","買賣超"] if x in rank.columns), None)
    sell_col = next((x for x in ["賣超張數","賣超","淨賣超"] if x in rank.columns), None)
    name_col = next((x for x in ["股票名稱","名稱","股票"] if x in rank.columns), None)
    code_col = next((x for x in ["股票代號","代號"] if x in rank.columns), None)
    if name_col and buy_col:
        rank[buy_col] = pd.to_numeric(rank[buy_col], errors="coerce").fillna(0)
        top_buy = rank.nlargest(10, buy_col)
        left, right = st.columns(2)
        with left:
            st.markdown("#### 🟢 買超前 10 名")
            cols = [x for x in [code_col,name_col,buy_col] if x]
            st.dataframe(top_buy[cols], hide_index=True, width="stretch")
        with right:
            st.markdown("#### 🔴 賣超前 10 名")
            if sell_col:
                rank[sell_col] = pd.to_numeric(rank[sell_col], errors="coerce").fillna(0)
                top_sell = rank.nlargest(10, sell_col)
            else:
                top_sell = rank.nsmallest(10, buy_col)
            cols = [x for x in [code_col,name_col,sell_col or buy_col] if x]
            st.dataframe(top_sell[cols], hide_index=True, width="stretch")
    else:
        st.info("排行 CSV 至少需要「股票名稱」與「買超／買超張數」欄位。")
else:
    st.caption("此區已預留在首頁最上方；接上可合法使用的排行資料來源後，可自動顯示每日買超／賣超前 10 名。")

with st.sidebar:
    st.header("輸入資料")
    stock = st.text_input("股票代號", "3189")
    window = st.selectbox("追蹤區間", [5, 10, 20, 60], index=2)
    method = st.selectbox("成本法", ["移動平均成本法"])
    uploaded = st.file_uploader("上傳分點每日資料 CSV", type=["csv"])
    st.markdown("CSV 欄位：日期、券商分點、買進張數、買進均價、賣出張數、賣出均價")

def moving_average(df):
    rows=[]
    for broker,g in df.sort_values("日期").groupby("券商分點"):
        qty=0.0; cost=0.0; total_buy=0.0; total_sell=0.0
        for _,r in g.iterrows():
            b=float(r["買進張數"]); bp=float(r["買進均價"]); s=float(r["賣出張數"])
            if b>0:
                cost=(qty*cost+b*bp)/(qty+b) if qty+b>0 else 0
                qty+=b; total_buy+=b
            if s>0:
                qty=max(0.0,qty-s); total_sell+=s
                if qty==0: cost=0.0
        rows.append({"券商分點":broker,"累積買進":total_buy,"累積賣出":total_sell,"估算剩餘張數":qty,"估算平均成本":cost})
    return pd.DataFrame(rows)

if uploaded:
    df=pd.read_csv(uploaded)
    need=["日期","券商分點","買進張數","買進均價","賣出張數","賣出均價"]
    missing=[x for x in need if x not in df.columns]
    if missing:
        st.error("缺少欄位："+"、".join(missing))
    else:
        df["日期"]=pd.to_datetime(df["日期"])
        max_date=df["日期"].max()
        trading_dates=sorted(df["日期"].drop_duplicates())
        keep_dates=trading_dates[-int(window):]
        df=df[df["日期"].isin(keep_dates)].copy()
        for x in need[2:]: df[x]=pd.to_numeric(df[x],errors="coerce").fillna(0)
        result=moving_average(df)
        current=st.number_input("目前股價（用於計算成本乖離）",min_value=0.0,value=0.0,step=0.5)
        if current>0:
            result["股價距成本(%)"]=result["估算平均成本"].apply(lambda x:(current/x-1)*100 if x>0 else 0)
        result["淨買超張數"]=result["累積買進"]-result["累積賣出"]
        result["動向"]=result["淨買超張數"].apply(lambda x:"🟢 加碼" if x>0 else ("🔴 減碼" if x<0 else "⚪ 持平"))
        if current > 0:
            result["交易訊號"] = result.apply(
                lambda r: (
                    "🟢 買入觀察" if r["估算平均成本"] > 0 and abs(current / r["估算平均成本"] - 1) <= 0.02 and r["淨買超張數"] > 0
                    else ("🔴 賣出觀察" if r["估算平均成本"] > 0 and current >= r["估算平均成本"] * 1.10 and r["淨買超張數"] < 0
                    else ("🟠 風險退出" if r["估算平均成本"] > 0 and current <= r["估算平均成本"] * 0.95 and r["淨買超張數"] < 0
                    else "⚪ 觀察"))
                ), axis=1
            )
        result=result.sort_values("估算剩餘張數",ascending=False)
        c1,c2,c3=st.columns(3)
        c1.metric("追蹤股票",stock)
        c2.metric(f"近 {window} 日追蹤分點",len(result))
        c3.metric("推估淨庫存",f'{result["估算剩餘張數"].sum():,.0f} 張')
        st.subheader(f"近 {window} 日分點成本排行")
        st.dataframe(result,hide_index=True,width="stretch")
        st.subheader("每日原始分點資料")
        st.dataframe(df.sort_values("日期",ascending=False),hide_index=True,width="stretch")
else:
    st.info("先上傳 CSV 即可開始計算。下一版可再接合法可用的自動資料來源。")
    sample=pd.DataFrame([
        ["2026-09-18","元大-某分點",100,420,20,430],
        ["2026-09-19","元大-某分點",50,440,30,450],
    ],columns=["日期","券商分點","買進張數","買進均價","賣出張數","賣出均價"])
    st.subheader("CSV 格式範例")
    st.dataframe(sample,hide_index=True,width="stretch")



with tab_foreign:
    st.header("🌍 外資券商買賣追蹤器")
    st.caption("從券商分點資料獨立追蹤外資券商，特別顯示摩根大通、美林與美商高盛。")
    foreign_file = st.file_uploader("上傳外資券商分點 CSV", type=["csv"], key="foreign_file")
    focus_names = ["摩根大通", "美林", "美林證券", "美商高盛", "美商高盛亞", "高盛"]
    if foreign_file:
        foreign = pd.read_csv(foreign_file)
        broker_col = next((x for x in ["券商分點","券商名稱","分點"] if x in foreign.columns), None)
        buy_col = next((x for x in ["買進張數","買張","買進"] if x in foreign.columns), None)
        sell_col = next((x for x in ["賣出張數","賣張","賣出"] if x in foreign.columns), None)
        price_col = next((x for x in ["均價","買進均價","買價"] if x in foreign.columns), None)
        stock_col = next((x for x in ["股票名稱","名稱","股票"] if x in foreign.columns), None)
        code_col = next((x for x in ["股票代號","代號"] if x in foreign.columns), None)
        if broker_col and buy_col and sell_col:
            foreign[buy_col] = pd.to_numeric(foreign[buy_col], errors="coerce").fillna(0)
            foreign[sell_col] = pd.to_numeric(foreign[sell_col], errors="coerce").fillna(0)
            foreign["淨買賣超"] = foreign[buy_col] - foreign[sell_col]
            foreign["重點外資"] = foreign[broker_col].astype(str).apply(
                lambda x: "摩根大通" if "摩根大通" in x else ("美林" if "美林" in x else ("高盛" if "高盛" in x else "其他外資"))
            )
            focus = foreign[foreign["重點外資"] != "其他外資"].copy()
            st.subheader("⭐ 摩根大通・美林・高盛")
            summary = focus.groupby("重點外資", as_index=False).agg(
                買進張數=(buy_col,"sum"), 賣出張數=(sell_col,"sum"), 淨買賣超=("淨買賣超","sum")
            )
            st.dataframe(summary, hide_index=True, width="stretch")
            if stock_col:
                st.subheader("重點外資個股進出")
                cols=[x for x in [code_col,stock_col,broker_col,buy_col,sell_col,"淨買賣超",price_col] if x]
                st.dataframe(focus.sort_values("淨買賣超",ascending=False)[cols], hide_index=True, width="stretch")
            st.subheader("全部外資券商資料")
            st.dataframe(foreign, hide_index=True, width="stretch")
        else:
            st.error("CSV 至少需要券商名稱、買進張數、賣出張數欄位。")
    else:
        st.info("外資券商分頁已建立。匯入分點資料後會自動抓出摩根大通、美林與高盛，計算買進、賣出與淨買賣超。")

@st.cache_data(ttl=900)
def fetch_taifex_institutional():
    """讀取期交所三大法人期貨未平倉；失敗時回傳空表，不讓整個 App 掛掉。"""
    url = "https://www.taifex.com.tw/cht/3/futContractsDate"
    headers = {"User-Agent": "Mozilla/5.0"}
    try:
        tables = pd.read_html(requests.get(url, headers=headers, timeout=12).text)
        frames = []
        for t in tables:
            flat = [" ".join(map(str, x)) if isinstance(x, tuple) else str(x) for x in t.columns]
            t.columns = flat
            joined = " ".join(flat) + " " + t.astype(str).to_string(index=False)
            if "外資" in joined and "未平倉" in joined:
                frames.append(t)
        return frames[0] if frames else pd.DataFrame()
    except Exception:
        return pd.DataFrame()

with tab_futures:
    st.header("📈 期貨市場")
    st.caption("追蹤台指期、微台、電子期、金融期與海外主要期貨；優先自動讀取期交所官方三大法人資料。")
    if st.button("🔄 更新期交所法人資料", key="refresh_taifex"):
        fetch_taifex_institutional.clear()
    auto_fut = fetch_taifex_institutional()
    if not auto_fut.empty:
        st.success("已取得期交所官方三大法人期貨資料")
        st.dataframe(auto_fut, hide_index=True, width="stretch")
    else:
        st.warning("目前未能自動取得期交所表格；仍可使用下方 CSV 備援，不影響其他分頁。")
    futures_file = st.file_uploader("上傳期貨行情／籌碼 CSV", type=["csv"], key="futures_file")
    if futures_file:
        fut = pd.read_csv(futures_file)
        st.dataframe(fut, hide_index=True, width="stretch")
        numeric_cols = fut.select_dtypes(include="number").columns.tolist()
        if "日期" in fut.columns:
            fut["日期"] = pd.to_datetime(fut["日期"], errors="coerce")
        st.subheader("期貨市場重點")
        wanted = [x for x in ["商品","契約","收盤價","漲跌幅","未平倉量","外資淨部位","投信淨部位","自營商淨部位","期現貨價差"] if x in fut.columns]
        if wanted:
            st.dataframe(fut[wanted], hide_index=True, width="stretch")
        if "日期" in fut.columns and "收盤價" in fut.columns:
            chart_df = fut.dropna(subset=["日期"]).set_index("日期")
            st.line_chart(chart_df["收盤價"])
    else:
        st.info("期貨分頁已建立。後續接入可合法使用的行情／法人未平倉資料後，可自動更新。")
        st.markdown("""
        **預計追蹤：** 台指期 TX、微型台指 MTX、電子期、金融期、台灣期交所三大法人未平倉、期現貨價差，以及美股指數／黃金等主要期貨。
        """)


with tab_pelosi:
    st.header("🏛️ Nancy / Paul Pelosi 公開披露持股與成本估算")
    st.caption("資料依美國眾議院 PTR 公開披露。交易金額多以區間申報，因此成本只能估算，並非精確成交成本。")
    st.link_button("開啟 Nancy Pelosi Stock Tracker", "https://nancypelosistocktracker.org/zh-TW")

    pelosi = pd.DataFrame([
        ["NVDA","NVIDIA","股票","Buy","2024-06-26",10000,1000000,5000000,None],
        ["NVDA","NVIDIA","股票","Buy","2024-07-26",10000,1000000,5000000,None],
        ["NVDA","NVIDIA","股票","Sell","2024-12-31",10000,1000000,5000000,None],
        ["AAPL","Apple","股票","Sell","2024-12-31",31600,5000000,25000000,None],
        ["AVGO","Broadcom","Call","Buy","2024-06-24",None,1000000,5000000,800],
        ["PANW","Palo Alto Networks","Call","Exercise","2024-12-20",None,1000000,5000000,100],
        ["NVDA","NVIDIA","Call","Exercise","2024-12-20",None,500000,1000000,12],
        ["GOOGL","Alphabet","Call","Buy","2025-01-14",None,250000,500000,150],
        ["AMZN","Amazon","Call","Buy","2025-01-14",None,250000,500000,150],
        ["VST","Vistra","Call","Buy","2025-01-14",None,500000,1000000,50],
        ["TEM","Tempus AI","Call","Buy","2025-01-14",None,50000,100000,20],
    ], columns=["代號","公司","類型","交易","日期","股數","金額下限","金額上限","履約價"])
    pelosi["日期"] = pd.to_datetime(pelosi["日期"])
    pelosi["申報金額中位數"] = (pelosi["金額下限"] + pelosi["金額上限"]) / 2
    pelosi["估算每股成本"] = pelosi.apply(
        lambda r: r["申報金額中位數"] / r["股數"] if pd.notna(r["股數"]) and r["股數"] > 0 and r["交易"] == "Buy"
        else (r["履約價"] if r["類型"] == "Call" and pd.notna(r["履約價"]) else None), axis=1
    )
    st.subheader("公開交易與成本估算")
    st.dataframe(pelosi, hide_index=True, width="stretch")

    st.subheader("成本估算方法")
    st.markdown("""
    - **已披露股票買進且有股數**：以「申報金額區間中位數 ÷ 股數」估算每股投入成本。
    - **選擇權**：履約價不是完整持股成本；完整成本還需要權利金、合約數與後續行權資訊，因此另外標示。
    - **申報金額為區間**：同時保留下限與上限，避免把中位數誤認為實際成交金額。
    - **披露有時間落差**：此頁呈現的是已公開申報資訊，不代表即時持倉。
    """)


with tab_mops:
    st.header("📢 台股重大訊息")
    st.caption("資料來源：臺灣證券交易所／櫃買中心公開資訊觀測站（MOPS）。")
    st.link_button("開啟公開資訊觀測站", "https://mops.twse.com.tw/mops/web/index")
    q = st.text_input("快速篩選股票代號／公司名稱／關鍵字", key="mops_query")
    st.info("此分頁已建立。下一步接入官方可用資料介面後，會在此顯示即時／當日重大訊息，並支援公司與關鍵字篩選。")
    st.markdown("""
    **預計特別標示的事件：**
    - 🔥 營收／獲利／財測重大變動
    - 💰 股利、庫藏股、現增／私募
    - 🏭 重大訂單、資本支出、投資與處分資產
    - 🤝 併購、策略合作、轉投資
    - ⚠️ 訴訟、處分、停工、災害、財務或交易異常
    - 👔 董事長／總經理等重要人事異動
    - 🗓️ 法說會、董事會與重大決議
    """)
