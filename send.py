import csv
from tencentcloud.common import credential
from tencentcloud.common.exception.tencent_cloud_sdk_exception import TencentCloudSDKException
from tencentcloud.sms.v20210111 import sms_client, models
from tencentcloud.common.profile.client_profile import ClientProfile
from tencentcloud.common.profile.http_profile import HttpProfile

def load_credentials_from_csv(csv_file="SecretKey.csv"):
    """从 CSV 文件加载腾讯云 SecretId 和 SecretKey"""
    try:
        with open(csv_file, mode="r") as file:
            reader = csv.DictReader(file)
            row = next(reader)  # 读取第一行
            return row["SecretId"], row["SecretKey"]
    except FileNotFoundError:
        raise Exception(f"CSV file '{csv_file}' not found. Please provide SecretId and SecretKey.")
    except KeyError:
        raise Exception("Invalid CSV format. Expected columns: 'SecretId', 'SecretKey'.")

try:
    # 从 CSV 加载密钥
    secret_id, secret_key = load_credentials_from_csv()
    cred = credential.Credential(secret_id, secret_key)

    # 实例化一个http选项，可选的，没有特殊需求可以跳过。
    httpProfile = HttpProfile()
    # 如果需要指定proxy访问接口，可以按照如下方式初始化hp（无需要直接忽略）
    # httpProfile = HttpProfile(proxy="http://用户名:密码@代理IP:代理端口")
    httpProfile.reqMethod = "POST"  # post请求(默认为post请求)
    httpProfile.reqTimeout = 10    # 请求超时时间，单位为秒(默认60秒)
    httpProfile.endpoint = "sms.tencentcloudapi.com"  # 指定接入地域域名(默认就近接入)

    # 非必要步骤:
    # 实例化一个客户端配置对象，可以指定超时时间等配置
    clientProfile = ClientProfile()
    clientProfile.signMethod = "TC3-HMAC-SHA256"  # 指定签名算法
    clientProfile.language = "en-US"
    clientProfile.httpProfile = httpProfile
    
    client = sms_client.SmsClient(cred, "ap-guangzhou", clientProfile)

    req = models.SendSmsRequest()

    req.SmsSdkAppId = "1400973406" # √
    # 短信签名内容: 使用 UTF-8 编码，必须填写已审核通过的签名
    # 签名信息可前往 [国内短信](https://console.cloud.tencent.com/smsv2/csms-sign) 或 [国际/港澳台短信](https://console.cloud.tencent.com/smsv2/isms-sign) 的签名管理查看
    req.SignName = "腾讯云"
    # 模板 ID: 必须填写已审核通过的模板 ID
    # 模板 ID 可前往 [国内短信](https://console.cloud.tencent.com/smsv2/csms-template) 或 [国际/港澳台短信](https://console.cloud.tencent.com/smsv2/isms-template) 的正文模板管理查看
    req.TemplateId = "449739"
    # 模板参数: 模板参数的个数需要与 TemplateId 对应模板的变量个数保持一致，，若无模板参数，则设置为空
    req.TemplateParamSet = ["1234"]
    # 下发手机号码，采用 E.164 标准，+[国家或地区码][手机号]
    # 示例如：+8613711112222， 其中前面有一个+号 ，86为国家码，13711112222为手机号，最多不要超过200个手机号
    req.PhoneNumberSet = ["+8615623169098", "+8615288361907"]
    # 用户的 session 内容（无需要可忽略）: 可以携带用户侧 ID 等上下文信息，server 会原样返回
    req.SessionContext = ""
    # 短信码号扩展号（无需要可忽略）: 默认未开通，如需开通请联系 [腾讯云短信小助手]
    req.ExtendCode = ""
    # 国内短信无需填写该项；国际/港澳台短信已申请独立 SenderId 需要填写该字段，默认使用公共 SenderId，无需填写该字段。注：月度使用量达到指定量级可申请独立 SenderId 使用，详情请联系 [腾讯云短信小助手](https://cloud.tencent.com/document/product/382/3773#.E6.8A.80.E6.9C.AF.E4.BA.A4.E6.B5.81)。
    req.SenderId = ""

    resp = client.SendSms(req)

    # 输出json格式的字符串回包
    print(resp.to_json_string(indent=2))

except TencentCloudSDKException as err:
    print(err)