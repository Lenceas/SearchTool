using System;
using TencentCloud.Common;
using TencentCloud.Common.Profile;
using TencentCloud.Ocr.V20181119;
using TencentCloud.Ocr.V20181119.Models;

namespace SearchTool.Test
{
    internal class Program
    {
        static void Main(string[] args)
        {
            try
            {
                Credential cred = new Credential
                {
                    SecretId = "",
                    SecretKey = ""
                };

                ClientProfile clientProfile = new ClientProfile();
                HttpProfile httpProfile = new HttpProfile();
                httpProfile.Endpoint = ("ocr.tencentcloudapi.com");
                clientProfile.HttpProfile = httpProfile;

                OcrClient client = new OcrClient(cred, "ap-guangzhou", clientProfile);
                GeneralAccurateOCRRequest req = new GeneralAccurateOCRRequest();
                req.ImageUrl = "F:\\小工具\\images\\loginbg.jpeg";
                req.ImageBase64 = "";
                GeneralAccurateOCRResponse resp = client.GeneralAccurateOCRSync(req);
                Console.WriteLine(AbstractModel.ToJsonString(resp));
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
            Console.Read();
        }
    }
}