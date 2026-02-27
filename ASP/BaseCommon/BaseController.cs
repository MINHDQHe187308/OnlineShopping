using System.Security.Cryptography;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Newtonsoft.Json;
using System.Net.Http.Headers;
using ASP.Models.Admin.Menus;

namespace ASP.BaseCommon
{
    public class BaseController : Controller
    {
        /// <summary>
        /// declare instance lib
        /// </summary>
        /// <param name="contextBase"></param>
        /// <param name="env"></param>
        public BaseController()
        {
        }
        
        /// <summary>
        /// Language option select list
        /// </summary>
        /// <returns></returns>
        public static List<SelectListItem> LangOptions()
        {
            var sltItem = new List<SelectListItem>();
            sltItem.Insert(0, new SelectListItem { Value = "en-US", Text = "English" });
            return sltItem;
        }

        /// <summary>
        /// handle sub string
        /// </summary>
        /// <param name="strM"></param>
        /// <param name="strP"></param>
        /// <param name="maxLength"></param>
        /// <returns></returns>
        public static string HandleSubString(string strM, string strP, int maxLength)
        {
            if (string.IsNullOrEmpty(strM))
            {
                // tra ve P
                if (!string.IsNullOrEmpty(strP))
                {
                    if (strP.Length > maxLength)
                        return strP.Substring(0, maxLength) + " ...";
                    else
                        return strP;
                }
                else
                {
                    return null;
                }

            }
            else
            {
                // tra ve M
                if (strM.Length > maxLength)
                    return strM.Substring(0, maxLength) + " ...";
                else
                    return strM;
            }
        }

        /// <summary>
        /// list static menu
        /// </summary>
        /// <returns></returns>
        public List<MenuDetail> GetMenuStatics()
        {
            var sltItem = new List<MenuDetail>();
            sltItem.Add(new MenuDetail { name = "Home", excerpt = "abc", thumbnail = "assets/image_404.png", url = "", target = "_self", languagekey = "0" });
            sltItem.Add(new MenuDetail { name = "Nhập phiếu", excerpt = "abc", thumbnail = "assets/image_404.png", url = "Inventory", target = "_self", languagekey = "0" });
            sltItem.Add(new MenuDetail { name = "PRO Check phiếu", excerpt = "abc", thumbnail = "assets/image_404.png", url = "Pro", target = "_self", languagekey = "0" });
            sltItem.Add(new MenuDetail { name = "ACC Check phiếu", excerpt = "abc", thumbnail = "assets/image_404.png", url = "Acc", target = "_self", languagekey = "0" });
            sltItem.Add(new MenuDetail { name = "PC duyệt phiếu", excerpt = "abc", thumbnail = "assets/image_404.png", url = "Pc", target = "_self", languagekey = "0" });
            sltItem.Add(new MenuDetail { name = "Inventory Tag", excerpt = "abc", thumbnail = "assets/image_404.png", url = "InventoryTag", target = "_self", languagekey = "0" });
            sltItem.Add(new MenuDetail { name = "Quản lý Tag", excerpt = "abc", thumbnail = "assets/image_404.png", url = "TagManagement", target = "_self", languagekey = "0" });
            sltItem.Add(new MenuDetail { name = "Kiểm kê kho", excerpt = "abc", thumbnail = "assets/image_404.png", url = "WarehouseInventory", target = "_self", languagekey = "0" });

            return sltItem;
        }

        /// <summary>
        /// message base with key
        /// </summary>
        /// <param name="key"></param>
        /// <returns></returns>
        public static string BaseMessage(string key)
        {
            var _dicMessage = new Dictionary<string, string>() {
                { "validator_fails"  ,  "Vui lòng kiểm tra lại dữ liệu"},
                { "create_success"  ,  "Tạo mới thành công"},
                { "create_fails"  ,  "Đã có lỗi xảy ra trong quá trình tạo mới, Vui lòng kiểm tra lại"},
                { "update_success"  ,  "Cập nhật thành công"},
                { "update_fails"  ,  "Đã có lỗi xảy ra trong quá trình cập nhật, Vui lòng kiểm tra lại"},
                { "banned_success"  ,  "Khóa thành công"},
                { "banned_fails"  ,  "Đã có lỗi xảy ra trong quá trình khóa dữ liệu, Vui lòng kiểm tra lại"},
                { "delete_success"  ,  "Xóa thành công"},
                { "delete_fails"  ,  "Đã có lỗi xảy ra trong quá trình xóa dữ liệu, Vui lòng kiểm tra lại"},
                { "reject_success"  ,  "Từ chối thành công"},
                { "reject_fails"  ,  "Đã có lỗi xảy ra trong quá trình cập nhật dữ liệu, Vui lòng kiểm tra lại"},
                { "row_fails"  ,  "Không tìm thấy bản ghi, hoặc bản ghi đã được xóa"},
            };
            return _dicMessage[key];
        }

        /// <summary>
        /// page size select
        /// </summary>
        /// <returns></returns>
        public static List<SelectListItem> BasePageSize()
        {
            var _list = new List<SelectListItem>()
            {
                new SelectListItem(){ Value = "10",Text = "10"},
                new SelectListItem(){ Value = "25",Text = "25"},
                new SelectListItem(){ Value = "50",Text = "50"},
                new SelectListItem(){ Value = "100",Text = "100"},
                new SelectListItem(){ Value = "500",Text = "500"},
            };
            return _list;
        }

        /// <summary>
        /// hash fuction for password
        /// </summary>
        /// <param name="password"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException"></exception>
        public static string HashPassword(string password)
        {
            byte[] salt;
            byte[] buffer2;
            if (password == null)
            {
                throw new ArgumentNullException("password");
            }
            using (Rfc2898DeriveBytes bytes = new Rfc2898DeriveBytes(password, 0x10, 0x3e8))
            {
                salt = bytes.Salt;
                buffer2 = bytes.GetBytes(0x20);
            }
            byte[] dst = new byte[0x31];
            Buffer.BlockCopy(salt, 0, dst, 1, 0x10);
            Buffer.BlockCopy(buffer2, 0, dst, 0x11, 0x20);
            return Convert.ToBase64String(dst);
        }

        /// <summary>
        /// return text level manage for html view
        /// </summary>
        /// <param name="val"></param>
        /// <returns></returns>
        public static string GetTextEnumLevelManage(int val = 0)
        {
            var sltEnum = Enum.GetValues(typeof(EnumLevelManage)).Cast<EnumLevelManage>()
            .Select(se => new SelectListItem
            {
                Text = se.ToString(),
                Value = ((int)se).ToString()
            }).FirstOrDefault(w => w.Value == val.ToString());
            var str = "<span class='badge badge-light'>" + sltEnum.Text + "</span>";
            if (val == 10)
            {
                str = "<span class='badge badge-success'>" + sltEnum.Text + "</span>";
            }
            else if (val == 15)
            {
                str = "<span class='badge badge-dark'>" + sltEnum.Text + "</span>";
            }
            return sltEnum != null ? str : "";
        }

        /// <summary>
        /// return text enum status user for html view
        /// </summary>
        /// <param name="val"></param>
        /// <returns></returns>
        public static string GetTextEnumStatusUser(int val = -5)
        {
            //
            var str = "<span class='badge badge-danger'>" + EnumStatusUser.InActive.GetDescription() + "</span>";
            if (val == 10)
            {
                str = "<span class='badge badge-success'>" + EnumStatusUser.Active.GetDescription() + "</span>";
            }
            else if (val == 5)
            {
                str = "<span class='badge badge-warning'>" + EnumStatusUser.Pending.GetDescription() + "</span>";
            }
            return str;
        }

        /// <summary>
        /// return text enum D.O.H type
        /// </summary>
        /// <param name="val"></param>
        /// <returns></returns>
        public static string GetDisplayUserType(int val = 0)
        {
            var _res = "";
            switch (val)
            {
                case (int)EnumUserType.PRO:
                    _res = $"<span class='badge badge-primary'>{EnumUserType.PRO.GetDescription()}</span>";
                    break;
                case (int)EnumUserType.ACC:
                    _res = $"<span class='badge badge-danger'>{EnumUserType.ACC.GetDescription()}</span>";
                    break;
                case (int)EnumUserType.PC:
                    _res = $"<span class='badge badge-success'>{EnumUserType.PC.GetDescription()}</span>";
                    break;
                case (int)EnumUserType.None:
                    _res = $"<span class='badge badge-secondary'>{EnumUserType.None.GetDescription()}</span>";
                    break;
                default:
                    break;
            }
            return _res;
        }

        public static string GetDisplayApproveLevel(short val = 0)
        {
            var _res = "";
            switch (val)
            {
                case (short)EnumLevelApprove.None:
                    _res = $"<span class='badge badge-secondary'>{EnumLevelApprove.None.GetDisplayName()}</span>"; break;
                case (short)EnumLevelApprove.Check:
                    _res = $"<span class='badge badge-success'>{EnumLevelApprove.Check.GetDisplayName()}</span>"; break;
                case (short)EnumLevelApprove.Approve:
                    _res = $"<span class='badge badge-primary'>{EnumLevelApprove.Approve.GetDisplayName()}</span>"; break;
                default:
                    break;
            }
            return _res;
        }

        public static string GetDisplayApproveStatus(short val = 0)
        {
            var _res = "";
            switch (val)
            {
                case (short)EnumAprroveStatus.Pending:
                    _res = $"<span class='badge badge-warning'>{EnumAprroveStatus.Pending.GetDisplayName()}</span>"; break;
                case (short)EnumAprroveStatus.Approved:
                    _res = $"<span class='badge badge-primary'>{EnumAprroveStatus.Approved.GetDisplayName()}</span>"; break;
                case (short)EnumAprroveStatus.Rejected:
                    _res = $"<span class='badge badge-success'>{EnumAprroveStatus.Rejected.GetDisplayName()}</span>"; break;
                default:
                    break;
            }
            return _res;
        }

        public static string GetDisplayAccCheckStatus(short val = 0)
        {
            var _res = "";
            switch (val)
            {
                case (short)EnumAccCheckStatus.OK:
                    _res = "<span class='badge badge-success'>OK</span>"; break;
                case (short)EnumAccCheckStatus.NG:
                    _res = "<span class='badge badge-warning'>NG</span>"; break;
                default:
                    break;
            }
            return _res;
        }

        /// <summary>
        /// return text enum product active
        /// </summary>
        /// <param name="val"></param>
        /// <returns></returns>
        public static string GetDisplayProductActive(int val = 0)
        {
            var _res = "";
            switch (val)
            {
                case (int)EnumStatusActive.Active:
                    _res = $"<span class='badge badge-primary'>{EnumStatusActive.Active.GetDescription()}</span>"; break;
                case (int)EnumStatusActive.Disable:
                    _res = $"<span class='badge badge-secondary'>{EnumStatusActive.Disable.GetDescription()}</span>"; break;
                default:
                    break;
            }
            return _res;
        }

        public static string GetDisplayLocation(int val = 0)
        {
            var _res = "";
            switch (val)
            {
                case (int)EnumLocationType.WIP:
                    _res = $"<span class='badge badge-primary'>{EnumLocationType.WIP.GetDescription()}</span>";
                    break;
                case (int)EnumLocationType.MIX:
                    _res = $"<span class='badge badge-warning'>{EnumLocationType.MIX.GetDescription()}</span>";
                    break;
                case (int)EnumLocationType.WH:
                    _res = $"<span class='badge badge-success'>{EnumLocationType.WH.GetDescription()}</span>";
                    break;
                default:
                    break;
            }
            return _res;
        }

        public static string GetDisplayUnitType(int val = 0)
        {
            var _res = "";
            switch (val)
            {
                case (int)EnumUnitType.EA:
                    _res = $"<span class='badge badge-primary'>{EnumUnitType.EA.GetDescription()}</span>";
                    break;
                case (int)EnumUnitType.GR:
                    _res = $"<span class='badge badge-secondary'>{EnumUnitType.GR.GetDescription()}</span>";
                    break;
                case (int)EnumUnitType.KG:
                    _res = $"<span class='badge badge-success'>{EnumUnitType.KG.GetDescription()}</span>";
                    break;
                case (int)EnumUnitType.MT:
                    _res = $"<span class='badge badge-danger'>{EnumUnitType.MT.GetDescription()}</span>";
                    break;
                case (int)EnumUnitType.ML:
                    _res = $"<span class='badge badge-warning'>{EnumUnitType.ML.GetDescription()}</span>";
                    break;
                case (int)EnumUnitType.LT:
                    _res = $"<span class='badge badge-info'>{EnumUnitType.LT.GetDescription()}</span>";
                    break;
                default:
                    break;
            }
            return _res;
        }

        public static string GetDisplayStatusRecall(bool status = false)
        {
            if (status)
                return "<span class='badge badge-success'>OK</span>";

            return "<span class='badge badge-warning'>NG</span>";
        }

        public static string GetDisplayRecall(bool status = false)
        {
            if (status)
                return "<span class='badge badge-success'>Đã thu hồi</span>";

            return "<span class='badge badge-warning'>Chưa thu hồi</span>";
        }

        /// <summary>
        /// return text status active
        /// </summary>
        /// <param name="val"></param>
        /// <returns></returns>
        public static string GetDisplayIsActive(bool val = false)
        {
            return ConvertBooleanVal(val, "Active", "Disable");
        }

        /// <summary>
        /// return text true false
        /// </summary>
        /// <param name="val"></param>
        /// <returns></returns>
        public static string GetDisplayTrueFalse(bool val = false)
        {
            return ConvertBooleanVal(val, "True", "False");
        }

        /// <summary>
        /// return text yes no
        /// </summary>
        /// <param name="val"></param>
        /// <returns></returns>
        public static string GetDisplayYesNo(bool val = false)
        {
            return ConvertBooleanVal(val, "Yes", "No");
        }

        /// <summary>
        /// return content api
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="strUrl"></param>
        /// <returns></returns>
        public IList<T> GetContentApi<T>(string strUrl = "")
        {
            #region get supplier
            HttpClientHandler clientHandler = new HttpClientHandler();
            clientHandler.ServerCertificateCustomValidationCallback = (sender, cert, chain, sslPolicyErrors) => { return true; };
            // Pass the handler to httpclient(from you are calling api)
            HttpClient client = new HttpClient(clientHandler);
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            var getContent = client.GetAsync(strUrl).Result;
            //
            var objsApi = new List<T>();
            if (getContent.IsSuccessStatusCode)
            {
                var resultApi = getContent.Content.ReadAsStringAsync();
                objsApi = JsonConvert.DeserializeObject<List<T>>(resultApi.Result);
            }
            //
            #endregion
            return objsApi;
        }

        /// <summary>
        /// get year
        /// </summary>
        /// <param name="fromY"></param>
        /// <returns></returns>
        public List<SelectListItem> GetYears(int fromY = 2022)
        {
            var list = new List<SelectListItem>();
            for (int i = fromY; i < DateTime.Now.Year + 4; i++)
            {
                list.Add(new SelectListItem { Value = i.ToString(), Text = i.ToString() });
            }
            return list;
        }

        /// <summary>
        /// get months
        /// </summary>
        /// <param name="en"></param>
        /// <returns></returns>
        public List<SelectListItem> GetMonths(bool en = false)
        {
            var list = new List<SelectListItem>();
            if (en)
            {
                foreach (var item in GetMonthReport())
                {
                    list.Add(new SelectListItem { Value = item.Value.ToString(), Text = item.Key });
                }
            }
            else
            {
                for (int i = 1; i <= 12; i++)
                {
                    list.Add(new SelectListItem { Value = i.ToString(), Text = i.ToString() });
                }
            }
            return list;
        }

        /// <summary>
        /// get month report
        /// </summary>
        /// <returns></returns>
        public Dictionary<string, int> GetMonthReport()
        {
            Dictionary<string, int> list = new Dictionary<string, int>();
            list.Add("Jan", 1);
            list.Add("Feb", 2);
            list.Add("Mar", 3);
            list.Add("Apr", 4);
            list.Add("May", 5);
            list.Add("Jun", 6);
            list.Add("Jul", 7);
            list.Add("Aug", 8);
            list.Add("Sep", 9);
            list.Add("Oct", 10);
            list.Add("Nov", 11);
            list.Add("Dec", 12);
            
            return list;
        }

        /// <summary>
        /// get month
        /// </summary>
        /// <returns></returns>
        public Dictionary<string, int> GetMonthReport2()
        {
            Dictionary<string, int> list = new Dictionary<string, int>();
            list.Add("Apr", 5);
            list.Add("May", 6);
            list.Add("Jun", 7);
            list.Add("Jul", 8);
            list.Add("Aug", 9);
            list.Add("Sep", 10);
            list.Add("Oct", 11);
            list.Add("Nov", 12);
            list.Add("Dec", 1);
            list.Add("Jan", 2);
            list.Add("Feb", 3);
            list.Add("Mar", 4);
            return list;
        }

        /// <summary>
        /// import policy factory, category
        /// </summary>
        /// <returns></returns>
        public Dictionary<string, int> GetMonthImportExcel()
        {
            Dictionary<string, int> list = new Dictionary<string, int>();
            list.Add("B", 4);
            list.Add("C", 5);
            list.Add("D", 6);
            list.Add("E", 7);
            list.Add("F", 8);
            list.Add("G", 9);
            list.Add("H", 10);
            list.Add("I", 11);
            list.Add("J", 12);
            list.Add("K", 1);
            list.Add("L", 2);
            list.Add("M", 3);
            return list;
        }
        
        /// <summary>
        /// export supplier
        /// </summary>
        /// <returns></returns>
        public Dictionary<int, string> GetMonthExportExcel()
        {
            Dictionary<int, string> list = new Dictionary<int, string>();
            list.Add(1, "B");
            list.Add(2, "C");
            list.Add(3, "D");
            list.Add(4, "E");
            list.Add(5, "F");
            list.Add(6, "G");
            list.Add(7, "H");
            list.Add(8, "I");
            list.Add(9, "J");
            list.Add(10, "K");
            list.Add(11, "L");
            list.Add(12, "M");
            list.Add(13, "N");
            return list;
        }

        /// <summary>
        /// active status select list
        /// </summary>
        /// <returns></returns>
        public List<SelectListItem> GetSltBooleanActive()
        {
            return ConvertBooleanSlt("Active", "Disable");
        }

        /// <summary>
        /// active select list
        /// </summary>
        /// <returns></returns>
        public List<SelectListItem> GetSltBooleanTrueFalse()
        {
            return ConvertBooleanSlt("True", "False");
        }

        /// <summary>
        /// yes no select list
        /// </summary>
        /// <returns></returns>
        public List<SelectListItem> GetSltBooleanYesNo()
        {
            return ConvertBooleanSlt("Yes", "No");
        }

        /// <summary>
        /// excel column select list
        /// </summary>
        /// <returns></returns>
        public List<SelectListItem> GetSltExelCol()
        {
            var _lst = new List<SelectListItem>();
            string[] characters = new string[] { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };

            foreach (string item in characters)
            {
                _lst.Add(new SelectListItem { Text = item, Value = item });
            }

            return _lst;
        }

        /// <summary>
        /// generate bool select list with values
        /// </summary>
        /// <param name="positiveVal"></param>
        /// <param name="negativeVal"></param>
        /// <returns></returns>
        private List<SelectListItem> ConvertBooleanSlt(string positiveVal, string negativeVal)
        {
            var _lst = new List<SelectListItem>();
            _lst.Add(new SelectListItem { Text = positiveVal, Value = "True" });
            _lst.Add(new SelectListItem { Text = negativeVal, Value = "False" });
            return _lst;
        }

        /// <summary>
        /// generate text bool with values
        /// </summary>
        /// <param name="val"></param>
        /// <param name="positiveVal"></param>
        /// <param name="negativeVal"></param>
        /// <returns></returns>
        private static string ConvertBooleanVal(bool val, string positiveVal, string negativeVal)
        {
            if (val) return $"<span class='badge badge-primary'>{positiveVal}</span>";
            return $"<span class='badge badge-secondary'>{negativeVal}</span>";
        }
    }
}
