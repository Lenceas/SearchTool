using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using TencentCloud.Common;
using TencentCloud.Common.Profile;
using TencentCloud.Ocr.V20181119;
using TencentCloud.Ocr.V20181119.Models;

namespace SearchTool
{
    /// <summary>
    /// 腾讯云OCR文字识别 每月免费额度 1000次 按顺序调用：
    /// 1、通用印刷体识别（高精度版）
    /// 2、通用印刷体识别
    /// 3、通用印刷体识别（精简版）
    /// 4、通用印刷体识别（高速版）
    /// </summary>
    public class EstateCertOCR
    {
        public static SecretModel InitSecret()
        {
            var result = new SecretModel();
            var txt = FileHelper.DifDBConnOfSecurity($@"{Application.StartupPath}secret_Id_Key.txt");
            if (!string.IsNullOrEmpty(txt))
            {
                var sp = txt.Split("-");
                if (sp.Length == 2)
                {
                    result.secretId = sp[0];
                    result.secretKey = sp[1];
                }
            }
            return result;
        }

        public class SecretModel
        {
            public SecretModel()
            {
                secretId = "";
                secretKey = "";
            }
            public string secretId { get; set; }
            public string secretKey { get; set; }
        }

        /// <summary>
        /// 图片文字识别
        /// </summary>
        /// <param name="imageFileName"></param>
        /// <param name="isImgBase64String"></param>
        /// <param name="imgBase64"></param>
        /// <returns></returns>
        public static string Ocr(string imageFileName, bool isImgBase64String = false, string imgBase64 = "")
        {
            var result = string.Empty;
            var ocrResModel = Ocr1(imageFileName, isImgBase64String, imgBase64);
            if (ocrResModel != null && ocrResModel.TextDetections.Count() > 0)
            {
                var keysList = ocrResModel?.TextDetections.Select(_ => _.DetectedText).Distinct() ?? new List<string>();
                result = string.Join(" ", keysList);
            }
            return result;
        }

        /// <summary>
        /// 图片文字识别 通用印刷体识别（高精度版）
        /// </summary>
        /// <param name="imageFileName">图片路径</param>
        /// <param name="isImgBase64String">是否识别已经转换好的Base64图片字符串,默认否</param>
        /// <param name="imgBase64">Base64图片字符串</param>
        /// <returns></returns>
        public static OcrResModel Ocr1(string imageFileName, bool isImgBase64String = false, string imgBase64 = "")
        {
            try
            {
                if (!isImgBase64String)
                {
                    if (string.IsNullOrEmpty(imageFileName))
                    {
                        throw new Exception("图片Url不能为空");
                    }

                    imgBase64 = ImgToBase64(imageFileName);
                    if (string.IsNullOrEmpty(imgBase64))
                    {
                        throw new Exception("图片转换Base64失败");
                    }
                }
                else
                {
                    if (string.IsNullOrEmpty(imgBase64))
                        throw new Exception("Base64图片字符串不能为空");
                }

                var secret = InitSecret();
                if (secret == null || string.IsNullOrEmpty(secret.secretId) || string.IsNullOrEmpty(secret.secretKey))
                {
                    throw new Exception("图别识别密钥加载失败");
                }
                Credential cred = new Credential { SecretId = secret.secretId, SecretKey = secret.secretKey };
                ClientProfile clientProfile = new ClientProfile();
                HttpProfile httpProfile = new HttpProfile();
                httpProfile.Endpoint = ("ocr.tencentcloudapi.com");
                clientProfile.HttpProfile = httpProfile;
                OcrClient client = new OcrClient(cred, "ap-guangzhou", clientProfile);
                GeneralAccurateOCRRequest req = new GeneralAccurateOCRRequest();
                //req.ImageUrl = "https://ocr-demo-1254418846.cos.ap-guangzhou.myqcloud.com/card/EstateCertOCR/EstateCertOCR1.jpg";
                req.ImageBase64 = imgBase64;
                GeneralAccurateOCRResponse resp = client.GeneralAccurateOCRSync(req);

                var result = AbstractModel.ToJsonString(resp);
                //json转为实体
                return JsonConvert.DeserializeObject<OcrResModel>(result);
            }
            catch (Exception e)
            {
                if (!string.IsNullOrEmpty(e.InnerException?.Message ?? string.Empty))
                {
                    var sp = e.InnerException?.Message.Split(":") ?? new string[0];
                    var code = sp?[1].Replace(" message", "") ?? string.Empty;
                    if (!string.IsNullOrEmpty(code) && code.Equals("ResourcesSoldOut.ChargeStatusException"))
                    {
                        return Ocr2(imageFileName);
                    }
                    else
                        throw new Exception(e.Message);
                }
                else
                    throw new Exception(e.Message);
            }
        }

        /// <summary>
        /// 图片文字识别 通用印刷体识别
        /// </summary>
        /// <param name="imageFileName">图片路径</param>
        /// <param name="isImgBase64String">是否识别已经转换好的Base64图片字符串,默认否</param>
        /// <param name="imgBase64">Base64图片字符串</param>
        /// <returns></returns>
        public static OcrResModel Ocr2(string imageFileName, bool isImgBase64String = false, string imgBase64 = "")
        {
            try
            {
                if (!isImgBase64String)
                {
                    if (string.IsNullOrEmpty(imageFileName))
                    {
                        throw new Exception("图片Url不能为空");
                    }

                    imgBase64 = ImgToBase64(imageFileName);
                    if (string.IsNullOrEmpty(imgBase64))
                    {
                        throw new Exception("图片转换Base64失败");
                    }
                }
                else
                {
                    if (string.IsNullOrEmpty(imgBase64))
                        throw new Exception("Base64图片字符串不能为空");
                }

                var secret = InitSecret();
                if (secret == null || string.IsNullOrEmpty(secret.secretId) || string.IsNullOrEmpty(secret.secretKey))
                {
                    throw new Exception("图别识别密钥加载失败");
                }
                Credential cred = new Credential { SecretId = secret.secretId, SecretKey = secret.secretKey };
                ClientProfile clientProfile = new ClientProfile();
                HttpProfile httpProfile = new HttpProfile();
                httpProfile.Endpoint = ("ocr.tencentcloudapi.com");
                clientProfile.HttpProfile = httpProfile;
                OcrClient client = new OcrClient(cred, "ap-guangzhou", clientProfile);
                GeneralBasicOCRRequest req = new GeneralBasicOCRRequest();
                //req.ImageUrl = "https://ocr-demo-1254418846.cos.ap-guangzhou.myqcloud.com/card/EstateCertOCR/EstateCertOCR1.jpg";
                req.ImageBase64 = imgBase64;
                GeneralBasicOCRResponse resp = client.GeneralBasicOCRSync(req);

                var result = AbstractModel.ToJsonString(resp);
                //json转为实体
                return JsonConvert.DeserializeObject<OcrResModel>(result);
            }
            catch (Exception e)
            {
                if (!string.IsNullOrEmpty(e.InnerException?.Message ?? string.Empty))
                {
                    var sp = e.InnerException?.Message.Split(":") ?? new string[0];
                    var code = sp?[1].Replace(" message", "") ?? string.Empty;
                    if (!string.IsNullOrEmpty(code) && code.Equals("ResourcesSoldOut.ChargeStatusException"))
                    {
                        return Ocr3(imageFileName);
                    }
                    else
                        throw new Exception(e.Message);
                }
                else
                    throw new Exception(e.Message);
            }
        }

        /// <summary>
        /// 图片文字识别 通用印刷体识别（精简版）
        /// </summary>
        /// <param name="imageFileName">图片路径</param>
        /// <param name="isImgBase64String">是否识别已经转换好的Base64图片字符串,默认否</param>
        /// <param name="imgBase64">Base64图片字符串</param>
        /// <returns></returns>
        public static OcrResModel Ocr3(string imageFileName, bool isImgBase64String = false, string imgBase64 = "")
        {
            try
            {
                if (!isImgBase64String)
                {
                    if (string.IsNullOrEmpty(imageFileName))
                    {
                        throw new Exception("图片Url不能为空");
                    }

                    imgBase64 = ImgToBase64(imageFileName);
                    if (string.IsNullOrEmpty(imgBase64))
                    {
                        throw new Exception("图片转换Base64失败");
                    }
                }
                else
                {
                    if (string.IsNullOrEmpty(imgBase64))
                        throw new Exception("Base64图片字符串不能为空");
                }

                var secret = InitSecret();
                if (secret == null || string.IsNullOrEmpty(secret.secretId) || string.IsNullOrEmpty(secret.secretKey))
                {
                    throw new Exception("图别识别密钥加载失败");
                }
                Credential cred = new Credential { SecretId = secret.secretId, SecretKey = secret.secretKey };
                ClientProfile clientProfile = new ClientProfile();
                HttpProfile httpProfile = new HttpProfile();
                httpProfile.Endpoint = ("ocr.tencentcloudapi.com");
                clientProfile.HttpProfile = httpProfile;
                OcrClient client = new OcrClient(cred, "ap-guangzhou", clientProfile);
                GeneralEfficientOCRRequest req = new GeneralEfficientOCRRequest();
                //req.ImageUrl = "https://ocr-demo-1254418846.cos.ap-guangzhou.myqcloud.com/card/EstateCertOCR/EstateCertOCR1.jpg";
                req.ImageBase64 = imgBase64;
                GeneralEfficientOCRResponse resp = client.GeneralEfficientOCRSync(req);

                var result = AbstractModel.ToJsonString(resp);
                //json转为实体
                return JsonConvert.DeserializeObject<OcrResModel>(result);
            }
            catch (Exception e)
            {
                if (!string.IsNullOrEmpty(e.InnerException?.Message ?? string.Empty))
                {
                    var sp = e.InnerException?.Message.Split(":") ?? new string[0];
                    var code = sp?[1].Replace(" message", "") ?? string.Empty;
                    if (!string.IsNullOrEmpty(code) && code.Equals("ResourcesSoldOut.ChargeStatusException"))
                    {
                        return Ocr4(imageFileName);
                    }
                    else
                        throw new Exception(e.Message);
                }
                else
                    throw new Exception(e.Message);
            }
        }

        /// <summary>
        /// 图片文字识别 通用印刷体识别（高速版）
        /// </summary>
        /// <param name="imageFileName">图片路径</param>
        /// <param name="isImgBase64String">是否识别已经转换好的Base64图片字符串,默认否</param>
        /// <param name="imgBase64">Base64图片字符串</param>
        /// <returns></returns>
        public static OcrResModel Ocr4(string imageFileName, bool isImgBase64String = false, string imgBase64 = "")
        {
            try
            {
                if (!isImgBase64String)
                {
                    if (string.IsNullOrEmpty(imageFileName))
                    {
                        throw new Exception("图片Url不能为空");
                    }

                    imgBase64 = ImgToBase64(imageFileName);
                    if (string.IsNullOrEmpty(imgBase64))
                    {
                        throw new Exception("图片转换Base64失败");
                    }
                }
                else
                {
                    if (string.IsNullOrEmpty(imgBase64))
                        throw new Exception("Base64图片字符串不能为空");
                }

                var secret = InitSecret();
                if (secret == null || string.IsNullOrEmpty(secret.secretId) || string.IsNullOrEmpty(secret.secretKey))
                {
                    throw new Exception("图别识别密钥加载失败");
                }
                Credential cred = new Credential { SecretId = secret.secretId, SecretKey = secret.secretKey };
                ClientProfile clientProfile = new ClientProfile();
                HttpProfile httpProfile = new HttpProfile();
                httpProfile.Endpoint = ("ocr.tencentcloudapi.com");
                clientProfile.HttpProfile = httpProfile;
                OcrClient client = new OcrClient(cred, "ap-guangzhou", clientProfile);
                GeneralFastOCRRequest req = new GeneralFastOCRRequest();
                //req.ImageUrl = "https://ocr-demo-1254418846.cos.ap-guangzhou.myqcloud.com/card/EstateCertOCR/EstateCertOCR1.jpg";
                req.ImageBase64 = imgBase64;
                GeneralFastOCRResponse resp = client.GeneralFastOCRSync(req);

                var result = AbstractModel.ToJsonString(resp);
                //json转为实体
                return JsonConvert.DeserializeObject<OcrResModel>(result);
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
        }

        /// <summary>
        /// 图片转化成base64字符串
        /// </summary>
        /// <param name="ImageFileName">图片路径</param>
        /// <returns></returns>
        public static string ImgToBase64(string ImageFileName)
        {
            try
            {
                Bitmap bmp = new Bitmap(ImageFileName);
                MemoryStream ms = new MemoryStream();
                bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                byte[] arr = new byte[ms.Length];
                ms.Position = 0;
                ms.Read(arr, 0, (int)ms.Length);
                ms.Close();
                return Convert.ToBase64String(arr);
            }
            catch (Exception)
            {
                return "";
            }
        }

        /// <summary>
        /// 图片转化成base64字符串
        /// </summary>
        /// <param name="bmp">图片</param>
        /// <returns></returns>
        public static string ImgToBase64(Bitmap bmp)
        {
            try
            {
                MemoryStream ms = new MemoryStream();
                bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                byte[] arr = new byte[ms.Length];
                ms.Position = 0;
                ms.Read(arr, 0, (int)ms.Length);
                ms.Close();
                return Convert.ToBase64String(arr);
            }
            catch (Exception)
            {
                return "";
            }
        }
    }
}
