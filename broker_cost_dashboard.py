import pandas as pd
import streamlit as st

st.set_page_config(page_title="券商分點成本追蹤", page_icon="🏦", layout="wide")
st.title("📊 籌碼與期貨市場追蹤儀表板")
tab_broker, tab_futures = st.tabs(["🏦 券商分點", "📈 期貨市場"])

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
    window = st.selectbox("追蹤區間", [5, 10, 20, 60], index=2)\n        method = st.selectbox("成本法", ["移動平均成本法"])
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
        df["日期"]=pd.to_datetime(df["日期"])\n            max_date=df["日期"].max()\n            trading_dates=sorted(df["日期"].drop_duplicates())\n            keep_dates=trading_dates[-int(window):]\n            df=df[df["日期"].isin(keep_dates)].copy()
        for x in need[2:]: df[x]=pd.to_numeric(df[x],errors="coerce").fillna(0)
        result=moving_average(df)
        current=st.number_input("目前股價（用於計算成本乖離）",min_value=0.0,value=0.0,step=0.5)
        if current>0:
            result["股價距成本(%)"]=result["估算平均成本"].apply(lambda x:(current/x-1)*100 if x>0 else 0)
        result["淨買超張數"]=result["累積買進"]-result["累積賣出"]\n            result["動向"]=result["淨買超張數"].apply(lambda x:"🟢 加碼" if x>0 else ("🔴 減碼" if x<0 else "⚪ 持平"))\n            result=result.sort_values("估算剩餘張數",ascending=False)
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


with tab_futures:
    st.header("📈 期貨市場")
    st.caption("追蹤台指期、微台、電子期、金融期與海外主要期貨；可匯入每日行情與法人籌碼資料。")
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
