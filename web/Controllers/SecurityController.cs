using MohwEmail.Filters;
using MohwEmail.Services;
using MohwEmail.ViewModels.Security;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using System.Xml;

namespace MohwEmail.Controllers
{
    [ErrorAttr]
    public class SecurityController : _BaseController
    {
        /// <summary>
        /// 登入
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        public ActionResult Login(string sys_id)
        {
            //            string soapBody =
            //@"<?xml version=""1.0"" encoding=""utf-8""?>
            //	<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">
            //	  <soap:Body>
            //		<rm:HelloWorld xmlns:rm=""http://tempuri.org/"">
            //		  <name>Rainmaker</name>
            //		</rm:HelloWorld>
            //	  </soap:Body>
            //	</soap:Envelope>";

            //HttpWebRequest req = (HttpWebRequest)WebRequest.Create("http://localhost:60591/WebService1.asmx");
            //req.Headers.Add("SOAPAction", "\"http://tempuri.org/HelloWorld\"");
            //req.ContentType = "text/xml;charset=\"utf-8\"";
            //req.Accept = "text/xml";
            //req.Method = "POST";
            //SSO
            if (sys_id != null)
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(Server.MapPath("/App_Data/ExcelTemplate/SSOxml.xml"));//載入xml檔
                XmlNode xn = xmlDoc.SelectSingleNode("//sys_id");
                string transDate = xn.InnerText;
                xn.InnerText = Guid.NewGuid().ToString();                
                xmlDoc.Save(Server.MapPath("/App_Data/ExcelTemplate/SSOxml.xml"));
                XmlDocument xmlDoc2 = new XmlDocument();
                xmlDoc2.Load(Server.MapPath("/App_Data/ExcelTemplate/SSOxml.xml"));//載入xml檔
                string sso_id = xmlDoc2.InnerText;
                //拋資料給SSO WebServers
                try
                {
                    string SendSOAPBody = @"<?xml version=""1.0"" encoding=""UTF-8""?>
                                      <SOAP-ENV:Envelope SOAP-ENV:encodingStyle = ""http://schemas.xmlsoap.org/soap/encoding/"" 
                                    xmlns:SOAP-ENV = ""http://schemas.xmlsoap.org/soap/envelope/"" 
                                    xmlns:xsd = ""http://www.w3.org/2001/XMLSchema""
                                    xmlns:xsi = ""http://www.w3.org/2001/XMLSchema-instance"" 
                                    xmlns:SOAP-ENC = ""http://schemas.xmlsoap.org/soap/encoding/""
                                    xmlns:si = ""http://soapinterop.org/xsd""
                                    xmlns:tns = ""urn:WellChoose"">                           
                                       <SOAP-ENV:Body>                             
                                       <tns:chkidno xmlns:tns = ""urn:WellChoose"">                              
                                       <sys_id xsi:type=""xsd:string"">"+sys_id+"</sys_id></tns:chkidno></SOAP-ENV:Body></SOAP-ENV:Envelope>";
                    using (StreamWriter tx = new StreamWriter(Server.MapPath("/App_Data/ExcelTemplate/login.txt"), true))
                    {
                        tx.WriteLine(SendSOAPBody);
                    }
                    //要改成webconfig參數
                    string _SSO_Path = ConfigurationManager.AppSettings["SSO_Path"];
                    HttpWebRequest req = (HttpWebRequest)WebRequest.Create(_SSO_Path);
                    //HttpWebRequest req = (HttpWebRequest)WebRequest.Create("http://203.65.102.118/soap/soapserver.php?wsdl");
                    req.Method = "POST";
                    req.ContentType = "text/xml";                  
                    using (Stream stm = req.GetRequestStream())
                    {
                        using (StreamWriter stmw = new StreamWriter(stm))
                        {
                            stmw.Write(SendSOAPBody);
                        }
                    }

                    WebResponse response = req.GetResponse();
                    using (StreamReader sr = new StreamReader(response.GetResponseStream()))
                    {
                        string responseBody = sr.ReadToEnd();
                        using (StreamWriter tx = new StreamWriter("/App_Data/ExcelTemplate/logOut.txt", true))
                        {
                            tx.WriteLine(responseBody);
                            XmlDocument doc = new XmlDocument();
                            doc.LoadXml(responseBody);//讀取回傳的String
                            XmlNode xn_id_ = doc.SelectSingleNode("//ID");
                            string ID_ = xn_id_.InnerText;
                            tx.WriteLine(ID_);
                            XmlNode xn_SUCCESS_ = doc.SelectSingleNode("//SUCCESS");
                            string SUCCESS_ = xn_SUCCESS_.InnerText;
                            tx.WriteLine(SUCCESS_);
                            ActionResult result1 = RedirectToAction("CaseQuery", "CaseManagement");
                            SecurityService securityService1 = new SecurityService();
                            try
                            {
                                base.user = securityService1.GetUser(ID_);
                                return result1;
                            }
                            catch (Exception ex)
                            {
                                tx.WriteLine(ex);
                                return RedirectToAction("Login");
                            }
                        }
                      
                        //XmlNamespaceManager mgr = new XmlNamespaceManager(doc.NameTable);
                        //mgr.AddNamespace("soap", "http://schemas.xmlsoap.org/soap/envelope/"); //這是SOAP 1.1
                        //var xmlNode = doc.SelectSingleNode("SOAP-ENV:Body", mgr);
                        //XmlDocument XmlDoc = new XmlDocument();
                        //XmlDoc.LoadXml(xmlNode.FirstChild.InnerText);
                        //XmlNodeList NodeLists = XmlDoc.SelectNodes("ns1:chkidnoResponse");
                        //foreach (XmlElement element in NodeLists)
                        //{
                        //    ID = element.GetElementsByTagName("ID")[0].InnerText;
                        //    EPNO = element.GetElementsByTagName("EPNO")[0].InnerText;
                        //    DPNO = element.GetElementsByTagName("DPNO")[0].InnerText;
                        //    DPNAME = element.GetElementsByTagName("DPNAME")[0].InnerText;
                        //    SUCCESS = element.GetElementsByTagName("SUCCESS")[0].InnerText;
                        //}
                    }
                }
                catch (Exception ex)
                {
                    using (StreamWriter tx = new StreamWriter("/App_Data/ExcelTemplate/logOut.txt", true))
                    {
                        tx.WriteLine(ex);
                    }
                        throw;
                }
              


                //XmlDocument xmlDoc3 = new XmlDocument();
                //xmlDoc3.Load(Server.MapPath("/App_Data/ExcelTemplate/LoginXml.xml"));//載入xml檔
                //XmlNode xn_id = xmlDoc3.SelectSingleNode("//ID");
                //string ID = xn_id.InnerText;
                //XmlNode xn_SUCCESS = xmlDoc3.SelectSingleNode("//SUCCESS");
                ////string EPNO = "";
                ////string DPNO = "";
                ////string DPNAME = "";
                //string SUCCESS = xn_SUCCESS.InnerText;
                //ActionResult result = RedirectToAction("CaseQuery", "CaseManagement");
                //SecurityService securityService = new SecurityService();
                //try
                //{
                //    base.user = securityService.GetUser("user01");
                //    return result;
                //}
                //catch (Exception)
                //{

                //    return RedirectToAction("Login");
                //}
            }

            //string ID = "";
            //string EPNO = "";
            //string DPNO = "";
            //string DPNAME = "";
            //string SUCCESS = "";
            //string SendSOAPBody = @"<?xml version=""1.0"" encoding=""UTF-8""?>
            //                  <SOAP-ENV:Envelope SOAP-ENV:encodingStyle = ""http://schemas.xmlsoap.org/soap/encoding/"" 
            //                xmlns:SOAP-ENV = ""http://schemas.xmlsoap.org/soap/envelope/"" 
            //                xmlns:xsd = ""http://www.w3.org/2001/XMLSchema""
            //                xmlns:xsi = ""http://www.w3.org/2001/XMLSchema-instance"" 
            //                xmlns:SOAP-ENC = ""http://schemas.xmlsoap.org/soap/encoding/""
            //                xmlns: si = ""http://soapinterop.org/xsd""
            //                xmlns: tns = ""urn:WellChoose"">                           
            //                   <SOAP-ENV:Body>                             
            //                   <tns:chkidno xmlns:tns = ""urn:WellChoose"">                              
            //                   <sys_id xsi:type = ""xsd:string"" > 37f6c07bb3047a0981f439b0f35a711b </sys_id>                                     
            //                                 </tns:chkidno >                                      
            //                              </SOAP-ENV:Body >
            //                             </SOAP-ENV:Envelope >
            //                                ";
            //HttpWebRequest req = (HttpWebRequest)WebRequest.Create("http://203.65.102.118/soap/soapserver.php?wsdl");
            //using (Stream stm = req.GetRequestStream())
            //{
            //    using (StreamWriter stmw = new StreamWriter(stm))
            //    {
            //        stmw.Write(SendSOAPBody);
            //    }
            //}

            //WebResponse response = req.GetResponse();

            //using (StreamReader sr = new StreamReader(response.GetResponseStream()))
            //{
            //    string responseBody = sr.ReadToEnd();
            //    XmlDocument doc = new XmlDocument();
            //    doc.LoadXml(responseBody);//讀取回傳的String
            //    XmlNamespaceManager mgr = new XmlNamespaceManager(doc.NameTable);
            //    mgr.AddNamespace("soap", "http://schemas.xmlsoap.org/soap/envelope/"); //這是SOAP 1.1
            //    var xmlNode = doc.SelectSingleNode("SOAP-ENV:Body", mgr);
            //    XmlDocument XmlDoc = new XmlDocument();
            //    XmlDoc.LoadXml(xmlNode.FirstChild.InnerText);
            //    XmlNodeList NodeLists = XmlDoc.SelectNodes("ns1:chkidnoResponse");
            //    foreach (XmlElement element in NodeLists)
            //    {
            //        ID = element.GetElementsByTagName("ID")[0].InnerText;
            //        EPNO = element.GetElementsByTagName("EPNO")[0].InnerText;
            //        DPNO = element.GetElementsByTagName("DPNO")[0].InnerText;
            //        DPNAME = element.GetElementsByTagName("DPNAME")[0].InnerText;
            //        SUCCESS = element.GetElementsByTagName("SUCCESS")[0].InnerText;
            //    }
            //}
            return View();
        }
        /// <summary>
        /// 執行登入
        /// </summary>
        /// <param name="loginViewModel"></param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult Login(LoginViewModel loginViewModel)
        {
            //判斷帳號登入是否為SSO
            //如果是透過sso登入 找sso的AD_VIEW_SSO帳號資訊
            //如果是部外人員透過帳號密碼登入 找Permission帳號資訊
            //部外人員驗證帳號密碼是否正確

            //兩種都要找帳號的角色清單因為只有Permission帳號資訊有角色資訊
            //User Models 需要有角色List & 帳號基本資料


            ActionResult result = RedirectToAction("CaseQuery", "CaseManagement");
            SecurityService securityService = new SecurityService();
            if ((loginViewModel.Message = securityService.ValidateUser(loginViewModel)).Equals(GlobalResource.Strings.LoginMessage.Success))
            {
                try
                {
                    base.user = securityService.GetUser(loginViewModel.UserId);
                }
                catch (Exception)
                {

                    return RedirectToAction("Login");
                }
                if (base.user.UserDetail.UserId.Contains("user"))
                {
                    return result;
                }
                else
                {
                    //if (base.user.UserDetail.Internal == "Y" )
                    //{
                    //    loginViewModel.Message = "驗證失敗 - " + "本部同仁請從員工入口網進入部長信箱後臺系統";
                    //    result = View(loginViewModel);
                    //    return result;
                    //}
                }
            }
            else
            {
                if (loginViewModel.Message.Contains("密碼錯誤"))
                {
                    ViewBag.Test = loginViewModel.Message;
                    loginViewModel.Message = "驗證失敗 - " + "密碼錯誤";
                }
                else
                {
                    loginViewModel.Message = "驗證失敗 - " + loginViewModel.Message;
                }
                result = View(loginViewModel);
                return result;
            }

            //登入紀錄
            //securityService.LogLogin(new LogLogin()
            //{
            //    IPAddress = Request.UserHostAddress,
            //    //IPAddress = securityService.GetClientIP(),
            //    IsSuccess = loginViewModel.Message.Equals(GlobalResource.Strings.LoginMessage.Success),
            //    LoginMessage = loginViewModel.Message,
            //    LogTime = DateTime.Now,
            //    InputUserId = loginViewModel.UserId,
            //    InputPassword = new SEC.TripleDES().Encrypt(loginViewModel.Password)
            //});

            return result;
        }

        public ActionResult UserMaintain()
        {
            return View();
        }
        /// <summary>
        /// 登出
        /// </summary>
        /// <returns></returns>
        public ActionResult Logout()
        {
            Session.Clear();
            return RedirectToAction("Login");
        }
    }
}