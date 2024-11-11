# Usage: streamlit run AIp05數據框轉換X.py --> http:/localhost:8501

import numpy as np
import pandas as pd
from st_aggrid import AgGrid #, GridUpdateMode, JsCode, ColumnsAutoSizeMode
import plotly.express as px
import plotly.graph_objs as go

import seaborn as sns
import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from datetime import date, datetime
from function import detail,sumX,age_use2,genPPTX,addBulletPage,addSlideDF,makeDFtable,create_ppt,aa


#%%####### (W).網站系統基本架構 ##########
import streamlit as st
from streamlit_navigation_bar import st_navbar

# 建筑年龄（百年老房、50年以上、20-50年、小于20年） 不同地区、每平均价、交易量、用户画像（商/居）
# 建筑材料（混凝土、木材、石材、其他） 每平均价、交易量、用户画像（商/居）
# 季度（四季度） 每平均价、交易量

#%%##===== (D4) 系統基本架構 [*網站.py(E4)] =====#####


#%%##===== (W1).自定公用函式庫 =====#####
# @st.cache_data
def getX(Xname):     ##== X = getX(Xname): 自 X.csv 讀取 X (KDD1), 並設定標籤 (KDD3) ==##
    a="1"
    if type(Xname)==type(a):
        
        XXX = pd.read_csv(Xname)
    else:
        XXX = Xname
    XXX["date"] = pd.to_datetime(XXX["datetime"]).dt.date
    XXX["year"] = pd.to_datetime(XXX["datetime"]).dt.year
    XXX["month"] = pd.to_datetime(XXX["datetime"]).dt.month
    XXX["yq"] = pd.PeriodIndex(XXX.date, freq='Q')
    # XXX["ym"] = pd.PeriodIndex(XXX.date, freq='M')

    XXX["quarter"] = pd.cut(XXX["month"], bins=[0, 3, 6, 9, 12], labels=["1", "2", "3", "4"])  #把月份离散处理成季度
    XXX["AR"] = pd.cut(XXX["age"],bins=[0,5,10,20,100], labels=["0~5", "6~10", "11~20", "20+"]) #屋齡离散化 AgeRange
    XXX["AR"] = XXX["AR"].cat.add_categories("缺失值")
    XXX["AR"] = XXX["AR"].fillna("缺失值")
    XXX["AR"] = pd.Categorical(XXX["AR"], categories=["0~5", "6~10", "11~20", "20+","缺失值"], ordered=True)
    XXX = XXX.sort_values("AR")

    
    return (XXX)



#%%##===== (W2).儀表板函式庫: 前台(a)navbar,(b)sidebar,(c)canvas,後台(d) =====#####
def 擷取交易(fname,data_explain):  ##== (KDD1)擷取交易儀表板: X = 擷取交易(fnameX) ==##
    ##== (d).後台 ==##
    X = getX(fname)
    print("\n\n>>>>> 擷取交易數據 (-->XXX) -----")  # -- 偵錯用


    data_explain = pd.DataFrame(data_explain)
    # 选择特定的列
    selected_columns = X[['year', 'month', 'yq',  'quarter', 'AR']]
    # 设置显示 index
    selected_columns_with_index = selected_columns.set_index(X.index)

    # 热力图
    corr_matrix = X.corr()
    plt.figure(figsize=(12, 8))
    sns.heatmap(corr_matrix, annot=True, cmap='coolwarm', center=0)

    # == (b).前台-sidebar ==##  #==> [[AIp04/C4)(2)垂直流程]]
    st.markdown("---")
    st.sidebar.header("== (KDD1) 擷取交易數據 -- ")
    st.sidebar.write("* 交易檔名 = " + fname)
    st.sidebar.write("* 記錄筆數 = ", X.shape[0])
    ##== (c).前台-canvas ==##


    title = '<h2 style="font-family:sans-serif;text-align:center">(KDD1) 擷取交易數據 %s</h2>'%(fname)
    st.markdown(title, unsafe_allow_html=True)
    st.subheader("1-1. 原始數據展示")
    st.write("* 數據來源：https://plvr.land.moi.gov.tw/DownloadOpenData")
    st.write("* 記錄筆數 = " + str(X.shape[0]))
    st.dataframe(X.head(5))
    # ==> [[AIp04/C4)(1)水平流程]]
    st.subheader("1-2. 數據說明")
    
    cols1 = st.columns([1,1])
    cols1[0].table(data_explain)
    st.markdown("---")

    
    st.html("<h2  style='text-align:center;margin:0 0 3% 0'>== (KDD2) 探索交易數據--</h2>")

    cols = st.columns([1, 1])  # -- (d).前台--canvas
    correlation_map = '<h3 style="text-align:center;font-family:sans-serif;">* Correlation</h3>'
    cols[0].markdown(correlation_map, unsafe_allow_html=True)
    cols[0].pyplot(plt)
    a = '<h3  style="text-align:center;font-family:sans-serif;">2.說明</h3>'
    cols[1].markdown(a, unsafe_allow_html=True)
    cols[1].html("""<p style='text-align: left;'>1. <span style="background-color:yellow">total_price & area：相關係數接近 0.94</span>，表現出<span style="background-color:yellow">強烈的正相關</span>。
                                                <br>這意味著<span style="background-color:yellow">房產總價（total_price）與面積（area） 之間關係密切</span>，面積越大，總價越高。
                                                <br>這是符合預期的。</p>
                    <p style='text-align: left;'>2. <span style="background-color:yellow">unit_price & age：相關係數約為 -0.43</span>，表現出中等強度的負相關。
                                                <br>說明房產單價（unit_price）隨著房齡（age）的增加而降低。<span style="background-color:yellow">房齡較大的房子單價可能相對較低。</span></p>
                    <p style='text-align: left;'>3. <span style="background-color:yellow">age & total_price</span>：相關係數為 -0.18，顯示出弱負相關，表示房齡較高的房產總價可能略低，但這種<span style="background-color:yellow">關係不強。</span></p>
                    <p style='text-align: left;'>4. <span style="background-color:yellow">unit_price & total_price</span>：相關係數為 0.34，顯示出弱到中等強度的正相關，意味著單價和總價之間存在一定程度的正相關關係，通常單價高的房產總價也較高，但<span style="background-color:yellow">影響程度不高。</span></p>
                    <p style='text-align: left;'>5. <span style="background-color:yellow">year、month 與其他變數</span>：年份和月份與其他變數的相關係數都非常小（接近 0），這意味著年份和月份對其他變數（例如總價、單價、面積等）幾乎沒有影響。 這可能是因為資料分佈在不同年份和月份的樣本變化不大，或者這些時間因素對價格和麵積等指標的<span style="background-color:yellow">影響並不顯著。</span></p>
                 """)   
    st.markdown("---")
    st.html("<h2  style='text-align:center;margin:0 0 3% 0'>== (KDD3) 交易數據轉換--</h2>")
    cols2 = st.columns([1, 1])
    cols2[0].subheader("1.產生的數據標籤")
    cols2[0].dataframe(selected_columns_with_index)
    cols2[1].subheader("2.說明")
    cols2[1].write('''
    * year：交易年份，由datetime轉化而來
    
    * month：交易月份，由datetime轉換而來
    
    * yq：交易年份和季度，由year&quarter轉換而來，2019Q1——2024Q1
    
    * quarter：交易季度，由month離散化而來
    
    * AR：屋齡範圍Age Range，由age離散化而來，分為 0-10,11-20,21-30,30-40 及 40年以上五個範圍
    ''')

    return X

def 季度模型(XXX):    ##== (KDD2)季度模型儀表板: Svyq = 總成交結構(X) ==##
    ##== (d).後台 ==##
    ##== (1).不同季度交易量

    Ta = pd.crosstab(XXX["地段"], XXX["yq"], margins=True);
    # start_period = pd.Period('2015Q1', freq='Q')
    ##== (2).頻次分布表
    Sv = pd.crosstab(index=XXX["地段"], columns=XXX["yq"],
                     values=XXX["unit_price"], aggfunc="mean", margins=True)


    FIGym = go.Figure(go.Scatter(x=Sv.T.index.astype(str), y=Sv.T["All"]))
    FIGym.update_traces(mode='markers+lines');
    FIGym.update_xaxes(tickformat='%Y-%m', dtick='M1');

    FIGym1 = px.pie(Ta[Ta.index != "All"], values="All", names=Ta[Ta.index != "All"].index, labels=Ta[Ta.index != "All"].index, color=Ta[Ta.index != "All"].index)


    fig = px.bar(Sv, x=Sv.index, y="All", color="All", text="All", barmode='group')

    fig1 = px.bar(Ta, x=Ta.index, y="All", color="All", text="All", barmode='group')

    # -- (B) 以下為 主畫面(canvas)設計
    st.markdown("---")
    st.header("== (KDD4) 交易模型（一）季度模型-- ")
    cols = st.columns([1, 1])  # -- (d).前台--canvas
    cols[0].subheader("1.1 四季度交易量")
    cols[0].dataframe(Ta)
    cols[1].subheader("1.2 交易量圓餅圖")
    cols[1].plotly_chart(FIGym1, theme="streamlit", use_container_width=True)
    st.subheader("1.3 數據解讀(KDD5) ")
    st.html(''' 
            <p>(1) <span style="background-color:yellow">忠孝東路、吳興街、永吉路、信義路和基隆路等核心街道屬於信義區的蛋黃區</span>，<span style="background-color:yellow">交易量明顯高於其他街道</span>，顯示出房地產需求集中於這些核心地段。這些街道擁有良好的交通接駁、商業設施和高生活機能，<span style="background-color:yellow">吸引大部分買家</span>。</p>
            <p>(2) <span style="background-color:yellow">松山路、光復南路等街道交易量雖不及核心街道，但仍有一定市場需求</span>，反映出買家對<span style="background-color:yellow">價格與地段的平衡考量</span>。</p>
            <p>(3) 位於<span style="background-color:yellow">邊緣的街道交易量較少</span>，顯示出需求較低，這些區域的購房主要源自預算考量或對特定社區的偏好，而非大眾市場需求。</p>
            <p>(4) 這張交易量圓餅圖展示信義區房市需求的地理分布，<span style="background-color:yellow">顯示購屋者偏好核心地帶，並逐漸向外擴展</span>，對未來房地產開發與市場定位具參考價值。</p>
                ''')
    st.markdown("---")
    cols1 = st.columns([1, 1])  # -- (d).前台--canvas
    cols1[0].subheader("2.1 四季度平均單價")
    cols1[0].dataframe(Sv)
    cols1[1].subheader("2.2 平均單價折線圖")
    cols1[1].plotly_chart(FIGym, theme="streamlit", use_container_width=True)
    st.subheader("2.3 數據解讀(KDD5)")
    st.html('''   
    <p>(1) <span style="background-color:yellow">2019年至2021年間平均單價在70至75之間微幅波動，總體穩定</span>，顯示當時房市需求相對穩定或供需平衡。<span style="background-color:yellow">2020年初疫情雖對市場有影響，但因經濟前景不確定，需求下降，價格未顯著上升</span>。</p>
    <p>(2) <span style="background-color:yellow">2021年疫情趨緩，需求快速釋放</span>，加上遠程辦公興起，富裕買家傾向在都市核心區購買更大或更高級的房產，<span style="background-color:yellow">房價因此上漲</span>。</p>
    <p>(3) <span style="background-color:yellow">2022上半年價格持續上漲</span>，因市場信心增強，買家預期房價會上漲，<span style="background-color:yellow">形成價格上漲的正向循環</span>；<span style="background-color:yellow">2022下半年價格大幅下跌</span>，可能因房市過熱和價格上漲過快，部分買家選擇觀望或退出市場。</p>
    <p>(4) <span style="background-color:yellow">2023上半年在疫情後需求延續「報復性消費」模式，價格持續上漲</span>；<span style="background-color:yellow">2023下半年受聯準會升息、地緣政治緊張、通膨等影響，價格下跌</span>。</p>    
             ''')

    return Sv

def 屋齡模型(XXX):

    AM = XXX.groupby("AR").agg({"地段": "nunique", "unit_price": "mean", "area":"count",
                                      }).reset_index()
    AM.columns = ["屋齡範圍","地段数","房單價","交易量"]

    # -- (B) 以下為 主畫面(canvas)設計
    st.markdown("---")
    st.header("== (KDD4) 交易模型（二）屋齡模型-- ")


    st.subheader("1. 屋齡交易量")
    st.dataframe(AM)

    st.subheader("2. 雷達圖")
    cols = st.columns([1, 1])  # -- (d).前台--canvas

    fig = px.funnel(AM, y="屋齡範圍", x="交易量", color="屋齡範圍", orientation="h")
    cols[0].subheader("2.1 不同屋齡範圍的交易量")
    cols[0].plotly_chart(fig, theme="streamlit", use_container_width=True)

    fig1 = px.funnel(AM, y="屋齡範圍", x="房單價", color="屋齡範圍", orientation="h")
    cols[1].subheader("2.2 不同屋齡範圍的房單價")
    cols[1].plotly_chart(fig1,theme="streamlit", use_container_width=True)
    st.subheader("3. 數據解讀(KDD5)")
    st.html('''
        <p>比較兩張圖:</p>
        <p>(1) <span style="background-color:yellow">屋齡對房屋購買有一定影響力</span>：新屋價格較高，但交易量低，顯示市場對價格敏感。隨著屋齡增加，<span style="background-color:yellow">價格下降，交易量增高</span>，許多買家偏好價格適中的房屋。</p>
        <p>(2) <span style="background-color:yellow">房屋價格隨屋齡增加而下降</span>，在<span style="background-color:yellow">(30,50)歲屋齡交易量達到高峰</span>，顯示部分買家願意購買屋齡較高的房屋以降低購房成本；<span style="background-color:yellow">50年以上屋齡的交易量下降</span>，反映市場對老屋需求低。</p>
        <p>(3) 無屋齡資料的房屋單價偏高，<span style="background-color:yellow">交易量排名第二</span>，可能位於特定地段或生活機能完善之處，即便無法判定屋齡，<span style="background-color:yellow">仍具高單價</span>。</p>
        <p>(4) 無屋齡資料的房屋單價偏高，推測是因為這些房屋包含了一些無法確認屋齡但價值較高的房產，或者位於特定地段的房屋，即便沒有屋齡資料，<span style="background-color:yellow">仍具有相對高的單價</span>。<p>
        <p>(5) <span style="background-color:yellow">屋齡越低單價越高</span>。</p>
    ''')

    return AM
    #AgeModel

# return places,number_df,unit_df,price_df
def select_option(df,places,usefor,ages):
    st.markdown("---")
    st.html("<h2>== (KDD4) 交易模型（三）路段金額--</h2>")
    option = st.selectbox("選擇一個選項", ['永吉路','信義路','基隆路','吳興街',
                                    '忠孝東路','虎林街','和平東路','嘉興街','松德路','松隆路','松仁路',
                                    '市民大道六段','福德街','富陽街','光復南路','松勤街','文昌街',
                                    '中坡南路','松平路','信安街','景雲街','松山路','紫雲街','崇德街',
                                    '東興路','松智路','中坡北路','莊敬路','松信路','仁愛路','松高路',
                                    '松勇路','逸仙路','大道路','林口街','青雲街','瑞雲街','祥雲街',
                                    '健康路'])
    st.write("您選擇了：", option)

    length , place_price,P,p = age_use2(df,places,ages,option,usefor)

    unit_df = sumX(df,ages,place_price/length,usefor)
    price_df = sumX(df,ages,place_price,usefor)
    score_df = sumX(df,ages,length,usefor)


    st.header(option+"段 每坪售出")
    P_heigh = '<p style="font-family:sans-serif; color:	SteelBlue; font-size: 20px;">最高 : %.2f 萬 ,用途 : %s ,屋齡 : %s 年 </p>'%(P[0],P[1],P[2])
    st.markdown(P_heigh, unsafe_allow_html=True)
    P_low = '<p style="font-family:sans-serif; color:SlateGray; font-size: 20px;">最低 : %.2f 萬 ,用途 : %s ,屋齡 : %s 年 </p>'%(p[0],p[1],p[2])
    st.markdown(P_low, unsafe_allow_html=True)
    
    year=[]
    #year1=[]
    for i in range(4):
        year.append([place_price[i][0],place_price[i][1],place_price[i][2],place_price[i][3],place_price[i][4]])
        #year0.append([length[i][0],length[i][1],length[i][2],length[i][3],length[i][4]])
    use2 = ['住家用', '商業用', '辦公用', '住商用', '工業用']

    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=use2,
        y=(year[0][0],year[0][1],year[0][2],year[0][3],year[0][4]),
        name='2~5',
        marker_color='indianred'
    ))
    fig.add_trace(go.Bar(
        x=use2,
        y=(year[1][0],year[1][1],year[1][2],year[1][3],year[1][4]),
        name='6~10',
        marker_color='lightsalmon'
    ))
    fig.add_trace(go.Bar(
        x=use2,
        y=(year[2][0],year[2][1],year[2][2],year[2][3],year[2][4]),
        name='11~20',
        marker_color='lightblue'
    ))
    fig.add_trace(go.Bar(
        x=use2,
        y=(year[3][0],year[3][1],year[3][2],year[3][3],year[3][4]),
        name='20+',
        marker_color='dodgerblue'
    ))
    

    # Here we modify the tickangle of the xaxis, resulting in rotated labels.
    fig.update_layout(barmode='group', xaxis_tickangle=-45)
    
    # 網頁上呈現
    ## 圖
    st.plotly_chart(fig)
    st.markdown("---")

    cols1 = st.columns([1,1])
    
    


    score_df.set_index('屋齡(年)', inplace=True)
    cols1[0].subheader("數量統計")
    cols1[0].write(score_df)
    
    unit_df.set_index('屋齡(年)', inplace=True)
    cols1[1].subheader("平均單一物件金額(單位: 萬)")
    cols1[1].write(unit_df.style.highlight_max(axis=0))
    
    
    price_df.set_index('屋齡(年)', inplace=True)
    cols1[0].subheader("金額統計(單位: 萬)")
    cols1[0].dataframe(price_df.style.highlight_max(axis=0),width=600)
    
    st.markdown("---")

    


    return places[option],score_df,unit_df,price_df,fig


def tmp(df,places,usefor,ages):
    a =1
    st.markdown("---")
    st.html("<h2>== (KDD4) 交易模型（四）單位金額 --</h2>")
    
    st.html("<h2>平均每坪售價</h2>")
    st.latex(r"""
             \frac{\sum_i^n unit\_price_i}{len(unit\_price)}
             """)
    st.html("<h3>以下單位均為:  坪/萬</h3>")

    cols2 = st.columns([1,1])
    cols3 = st.columns([1,1])
    cols4 = st.columns([1,1,1])
    a0,a1,a2,a3,a4 = detail(df,places,ages)
    p = aa(a0,ages)
    cols2[0].html("<h4>用途 : 住家用</h4>")
    cols2[0].text("")
    cols2[0].write(p)
    cols2[0].dataframe(a0.style.highlight_max())
    p = aa(a1,ages)
    cols2[1].html("<h4>商業用</h4>")
    cols2[1].write(p)
    cols2[1].dataframe(a1.style.highlight_max())

    p = aa(a2,ages)
    cols3[0].html("<h4>辦公用</h4>")
    cols3[0].write(p)
    cols3[0].dataframe(a2.style.highlight_max())
    p = aa(a3,ages)
    cols3[1].html("<h4>住商用</h4>")
    cols3[1].write(p)
    cols3[1].dataframe(a3.style.highlight_max())
    
    p = aa(a4,ages)
    cols4[0].html("<h4>工業用</h4>")
    cols4[0].write(p)
    cols4[0].dataframe(a4.style.highlight_max())

    
    return a0,a1,a2,a3,a4


def result(raw,df,data_explain,ppt_name,places,use,ages):
    st.html("<h1>匯出ppt      2019-2024 信義區房價分析</h1>")
    translate_raw = getX(raw)
    
    prs = create_ppt(raw,df,data_explain,ppt_name,places,use,translate_raw,ages)
    prs.save("ppt/test1.pptx")   #== (3).存檔 
    if True:
        st.html("<h1>success !</h1>")
    return 

#%%##===== (W3).導航函式庫 =====#####
def check2log(textStr,log):      ##== check 再將 textStr 納入 log 中, 並中並可以提供建議 ==##
    st.sidebar.markdown('---')
    st.sidebar.header("== (KDD5)請用戶輸入解讀--")
    st.session_state.username = st.sidebar.text_input("", st.session_state.username, placeholder="輸入用戶名")
    sugg_key = f"log_checkbox_{len(log)}"
    suggestion = st.sidebar.text_area("", value=st.session_state.suggestion, key=sugg_key, height=20, placeholder="輸入建議")
    if st.sidebar.button("LOG操作 / 提交建議"):
        log.append(textStr)
        if st.session_state.username and suggestion:
            log.append(f"<{st.session_state.username}建議>> {suggestion}")
            st.sidebar.success("建议已提交并纳入LOG");     st.session_state.suggestion = ""
        else:
            st.sidebar.error("不列入建議,只單純LOG操作內容")
    return

def initSSS(variables, pjName):  ##== 初始化 state_session 的各變量 ==##
    sss = st.session_state
    if "LOG" not in sss:
        sss.LOG = [pjName]   #-- 初始化 LOG 列表
    if "username" not in sss:
        sss.username = ""    #-- 初始化 username 列表
    if "suggestion" not in sss:
        sss.suggestion = ""  #-- 初始化 suggestion 列表
    if "prs" not in sss:
        sss.prs = Presentation();     sss.prs = genPPTX(pjName, "date: " + str(date.today()))
    for var in variables:       #-- 初始化传入的变量名为 None
        if var not in sss: sss[var] = None
    return sss


#%%##===== (W4).網站架構 =====#####
if __name__ == "__main__":
    
    ##== (1).設定頁面組態 與 導航列 (前台(a)navbar) ==##
    st.set_page_config(page_title="SPC-S01 RDS系統", page_icon="✅", layout="wide",)  #==> [[AIp04/C4)(5)加上頁註,頁標題等]]
    # st.set_option('deprecation.showPyplotGlobalUse', False)
    page = st_navbar(["[擷取交易]", "[季度模型]","[屋齡模型]","[路段選擇]","[單位金額]","[匯出PPT檔]"])

    ##== (2).設定session初始值等 ==##
    Xname = "data/Xinyi.csv"
    df="data/clean5.csv"
    sss = initSSS(["X", "TWH", "Svyq", "Xname","df"], "AIp03圖形可視化W"+"--"+Xname)
    sss.Xname = Xname
    sss.df = df
    df = pd.read_csv("data/clean5.csv")
    ##== (3).設定 前台((b)sidebar + (c)canvas)主標題 ==##
    title = '<h1 style="font-family:sans-serif;text-align:center;margin: 0 0 5% 0;">2019-2024 信義區房價分析</h1>'
    st.markdown(title, unsafe_allow_html=True)
    
    st.sidebar.title("初步運營分析(S01)控制盤--")


    places = {'永吉路':0,'信義路':1,'基隆路':2,'吳興街':3,'忠孝東路':4,'虎林街':5,'和平東路':6,'嘉興街':7,
              '松德路':8,'松隆路':9,'松仁路':10 ,'市民大道六段':11,'福德街':12,'富陽街':13,'光復南路':14,
              '松勤街':15,'文昌街':16,'中坡南路':17,'松平路':18,'信安街':19,'景雲街':20,'松山路':21,
              '紫雲街':22,'崇德街':23,'東興路':24,'松智路':25,'中坡北路':26,'莊敬路':27,'松信路':28,
              '仁愛路':29,'松高路':30,'松勇路':31,'逸仙路':32,'大道路':33,'林口街':34,'青雲街':35,
              '瑞雲街':36,'祥雲街':37,'健康路':38}
    usefor = {0:'住家用', 1:'商業用', 2:'辦公用', 3:'住商用', 4:'工業用'}
    ages = ["2~5", "6~10","11~20","20+"]
    data_explain = {
        "columns": [
            "地段","datetime","total_price","unit_price","area","主建物佔比","型態",
            "age","樓別/樓高","交易標的","交易筆棟數","建物現況格局","parking_price","管理組織",
            "電梯","主要用途","備註"
        ],
        "說明": [
            "信義區的路段","交易日期","總價（以台幣為單位）", "單位價格", "坪數","實際可使用坪數","房子型態",
            "屋齡","樓別/樓高","實際獲得","--","--","車位價格","有無管理單位",
            "有無電梯","住家or商等","--"
        ]
    }
    ##== (4).導航切換: 前台(a)navbar-->儀表板函式(b,c,d) ==##
    match page:
        case "[擷取交易]":
            sss.X = 擷取交易(sss.Xname,data_explain)
            check2log(f"擷取交易: {sss.Xname} to get X with {sss.X.shape[0]} records", sss.LOG)
        case "[季度模型]":
            if sss.X is None:
                st.write("尚未擷取交易數據，請先擷取交易數據！")
                sss.LOG.append("尚未擷取交易數據，請先擷取交易數據！")
            else:
                sss.Svyq = 季度模型(sss.X)
                check2log(f"季度模型: Svyq with {sss.Svyq.shape} shape", sss.LOG)
        case "[屋齡模型]":
            if sss.X is None:
                st.write("尚未擷取交易數據，請先擷取交易數據！")
                sss.LOG.append("尚未擷取交易數據，請先擷取交易數據！")
            else:
                sss.Svyq = 屋齡模型(sss.X)
                check2log(f"屋齡模型: Svyq with {sss.Svyq.shape} shape", sss.LOG)

        case "[路段選擇]":
            if sss.df is None:
                st.write("尚未擷取交易數據，請先擷取交易數據！")
                sss.LOG.append("尚未擷取交易數據，請先擷取交易數據！")
            else:
                sss.Svyq = select_option(df,places,usefor,ages)
                check2log(f"路段選擇: Svyq with {sss.Svyq[1].shape} shape", sss.LOG)
        case "[單位金額]":
            if sss.df is None:
                st.write("尚未擷取交易數據，請先擷取交易數據！")
                sss.LOG.append("尚未擷取交易數據，請先擷取交易數據！")
            else:
                sss.Svyq = tmp(df,places,usefor,ages)
                check2log(f"路段選擇: Svyq with {sss.Svyq[1].shape} shape", sss.LOG)

        case "[匯出PPT檔]":
            if sss.df is None:
                st.write("尚未擷取交易數據，請先擷取交易數據！")
                sss.LOG.append("尚未擷取交易數據，請先擷取交易數據！")
            else:
                raw = pd.read_csv("data/Xinyi.csv")
                use = ["地段","datetime","total_price","unit_price","area","age","主要用途"]
                raw = raw[["地段","datetime","total_price","unit_price","area","age","主要用途"]]
                result(raw.head(2),df,data_explain,"2019-2024 信義區房價分析",places,use,ages)

    ##== (5).操作日誌 ==##
    st.sidebar.markdown('<h2 style="color: blue;">操作LOG日誌</h2>', unsafe_allow_html=True)
    for i, log in enumerate(sss.LOG, 1): st.sidebar.write(f"({i}). {log}")
