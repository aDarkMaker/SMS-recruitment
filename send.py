cred = credential.Credential("TENCENTCLOUD_SECRET_ID", "TENCENTCLOUD_SECRET_KEY") 
# -*- coding: utf-8 -*-
import pandas as pd
from tencentcloud.common import credential
from tencentcloud.common.exception.tencent_cloud_sdk_exception import TencentCloudSDKException
from tencentcloud.sms.v20210111 import sms_client, models
from tencentcloud.common.profile.client_profile import ClientProfile
from tencentcloud.common.profile.http_profile import HttpProfile

def send_sms():
    try:
        # 1. 认证信息配置（请替换为您的实际密钥）
        cred = credential.Credential(
            "TENCENTCLOUD_SECRET_ID",
            "TENCENTCLOUD_SECRET_KEY"
        )

        # 2. 客户端配置
        httpProfile = HttpProfile()
        httpProfile.endpoint = "sms.tencentcloudapi.com"
        clientProfile = ClientProfile()
        clientProfile.httpProfile = httpProfile
        client = sms_client.SmsClient(cred, "ap-guangzhou", clientProfile)

        # 3. 读取Excel数据
        excel_path = r"path_to_your_excel_file"  # 使用原始字符串处理路径
        df = pd.read_excel(excel_path)

        # 4. 遍历Excel每一行
        for index, row in df.iterrows():
            try:
                # 提取数据
                name = str(row["名字"]) 
                raw_phone = str(row["电话"]).strip() 
                date = str(row["日期"])  # 格式为YYYY.MM.DD
                time = str(row["面试时间"])
                place = str(row["面试地点"])

                # 处理电话号码格式
                phone = "+86" + raw_phone if not raw_phone.startswith("+") else raw_phone

                # 5. 构建模板参数（根据实际模板调整参数顺序）
                # 假设模板内容为："{1}您好，您的面试时间为{2} {3}，地点为{4}"
                template_params = [name, date, time, place]
                # print(f"发送给 {name} 的参数：{template_params + [phone]}") # 调试

                # 6. 创建请求对象
                req = models.SendSmsRequest()
                req.SmsSdkAppId = "xxxxxxxx"  # 实际SDK的AppID
                req.SignName = "xxx"         # 已审核的签名
                req.TemplateId = "xxxxxx"       # 已审核的模板ID
                req.TemplateParamSet = template_params
                req.PhoneNumberSet = [phone]   

                # 7. 发送短信
                resp = client.SendSms(req)
                print(f"发送结果给 {name}：{resp.to_json_string(indent=2)}")

            except Exception as e:
                print(f"第 {index+1} 行处理失败，错误信息：{str(e)}")
                continue  # 跳过当前错误继续执行

    except TencentCloudSDKException as err:
        print(f"腾讯云API错误：{err}")
    except Exception as e:
        print(f"程序运行错误：{str(e)}")

if __name__ == "__main__":
    send_sms()