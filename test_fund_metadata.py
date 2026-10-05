from fund_metadata import _resolve_moneydj_wrapper_url, extract_fund_metadata


def test_moneydj_wrapper_and_metadata():
    wrapper = (
        "https://tcbbankfund.moneydj.com/main.html?"
        "sUrl=%24W%24WR%24WR03%5DDJHTM%7BA%7DACPS10-5808"
    )
    assert _resolve_moneydj_wrapper_url(wrapper) == (
        "https://tcbbankfund.moneydj.com/w/wr/wr03.djhtm?a=ACPS10-5808"
    )
    html = (
        "<h1>國內基金-績效比較</h1>"
        "<h3>5808統一奔騰基金 基金</h3>"
        "<div>指標指數 台灣資訊科技指數</div>"
    )
    result = extract_fund_metadata(html)
    assert result["fund_name"] == "5808統一奔騰基金"
    assert result["benchmark"] == "台灣資訊科技指數"
