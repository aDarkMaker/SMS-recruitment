# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
from tencentcloud.common import credential
from tencentcloud.common.exception.tencent_cloud_sdk_exception import TencentCloudSDKException
from tencentcloud.sms.v20210111 import sms_client, models
from tencentcloud.common.profile.client_profile import ClientProfile
from tencentcloud.common.profile.http_profile import HttpProfile
import io

# 页面基础配置
st.set_page_config(
    page_title="招新短信发送助手",
    page_icon="📨",
    layout="wide"
)

# 初始化Session State
if 'sent_records' not in st.session_state:
    st.session_state.sent_records = []

def send_sms(df, template_params):
    """发送短信的核心函数"""
    try:
        # 认证信息配置（建议通过Secrets管理敏感信息）
        cred = credential.Credential(
            st.secrets["TENCENT"]["SECRET_ID"],   # 从Streamlit Secrets获取
            st.secrets["TENCENT"]["SECRET_KEY"]   # 生产环境推荐使用加密方案
        )

        # 客户端配置
        httpProfile = HttpProfile()
        httpProfile.endpoint = "sms.tencentcloudapi.com"
        clientProfile = ClientProfile()
        clientProfile.httpProfile = httpProfile
        client = sms_client.SmsClient(cred, "ap-guangzhou", clientProfile)

        # 遍历DataFrame
        progress_bar = st.progress(0)
        total_rows = len(df)
        results = []

        for index, row in df.iterrows():
            try:
                # 进度更新
                progress = (index + 1) / total_rows
                progress_bar.progress(progress)

                # 数据提取
                name = str(row["名字"])
                raw_phone = str(row["电话"]).strip()
                date = str(row["日期"])
                time = str(row["面试时间"])
                place = str(row["面试地点"])

                # 电话号码处理
                phone = "+86" + raw_phone if not raw_phone.startswith("+") else raw_phone

                # 构建请求
                req = models.SendSmsRequest()
                req.SmsSdkAppId = st.secrets["TENCENT"]["APP_ID"]
                req.SignName = st.secrets["TENCENT"]["SIGN_NAME"]
                req.TemplateId = st.secrets["TENCENT"]["TEMPLATE_ID"]
                req.TemplateParamSet = [name, date, time, place]
                req.PhoneNumberSet = [phone]

                # 发送短信
                resp = client.SendSms(req)
                
                # 记录结果
                result = {
                    "姓名": name,
                    "电话": phone,
                    "状态": "成功" if resp.SendStatusSet[0].Code == "Ok" else "失败",
                    "消息": resp.SendStatusSet[0].Message,
                    "时间": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")
                }
                results.append(result)

            except Exception as e:
                results.append({
                    "姓名": name,
                    "电话": phone,
                    "状态": "异常",
                    "消息": str(e),
                    "时间": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")
                })
                continue

        return pd.DataFrame(results)

    except TencentCloudSDKException as err:
        st.error(f"腾讯云API错误：{err}")
    except Exception as e:
        st.error(f"程序运行错误：{str(e)}")
    finally:
        progress_bar.empty()

def preview_message(template, row):
    """生成短信预览内容"""
    try:
        return template.format(
            name=row["名字"],
            date=row["日期"],
            time=row["面试时间"],
            place=row["面试地点"]
        )
    except KeyError as e:
        st.error(f"模板参数错误，缺少字段：{e}")
        return None

# 页面布局
st.title("📨 招新短信发送助手")
st.markdown("---")

# 侧边栏配置
with st.sidebar:
    st.header("配置参数")
    
    # 文件上传组件
    uploaded_file = st.file_uploader(
        "上传Excel文件",
        type=["xlsx", "xls"],
        help="请确保文件包含以下列：名字、电话、日期、面试时间、面试地点"
    )
    
    # 短信模板编辑
    template = st.text_area(
        "短信模板内容",
        value="【招生宣传联合会】亲爱的{name}同学：我们是华中科技大学招生宣传联合会，感谢你选择招生宣传联合会这个大家庭，第一轮面试将在{date}的{time}于{place}进行，预计二十分钟，请提前十分钟到场签到。我们期待你的精彩表现！收到请回复“姓名+是否能参加面试”，若无法按时到场参加面试请说明原因，我们将为你调整面试时间，谢谢！",
        height= 270,
        help="谨记，修改模版不会改变短信内容，仅供预览"
    )
    
    # 发送按钮
    send_button = st.button("开始发送", type="primary")

# 主内容区域
if uploaded_file:
    try:
        # 读取Excel文件
        df = pd.read_excel(uploaded_file)
        
        # 显示数据预览
        st.subheader("数据预览")
        st.dataframe(df.head(3), use_container_width=True)
        
        # 显示短信预览
        st.subheader("短信内容预览")
        sample_row = df.iloc[0]
        preview = preview_message(template, sample_row)
        if preview:
            st.code(preview, language="text")

        # 发送流程
        if send_button:
            st.markdown("---")
            st.subheader("发送进度")
            
            # 执行发送
            results = send_sms(df, template)
            
            # 显示发送结果
            if results is not None:
                st.session_state.sent_records = results
                st.subheader("发送结果统计")
                col1, col2, col3 = st.columns(3)
                col1.metric("总发送量", len(results))
                col2.metric("成功数", len(results[results["状态"] == "成功"]))
                col3.metric("失败数", len(results[results["状态"] == "失败"]))
                
                st.dataframe(
                    results,
                    use_container_width=True,
                    column_config={
                        "时间": st.column_config.DatetimeColumn(
                            "发送时间",
                            format="YYYY-MM-DD HH:mm:ss"
                        )
                    }
                )
                
                # 添加下载按钮
                csv = results.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="下载发送记录",
                    data=csv,
                    file_name='sms_results.csv',
                    mime='text/csv'
                )
    except Exception as e:
        st.error(f"文件处理错误：{str(e)}")
else:
    st.info("👈 请在左侧上传Excel文件开始操作")

# 显示历史记录
if st.session_state.sent_records:
    st.markdown("---")
    st.subheader("历史发送记录")
    st.dataframe(st.session_state.sent_records, use_container_width=True)