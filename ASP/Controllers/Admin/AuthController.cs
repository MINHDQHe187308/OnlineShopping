using ASP.BaseCommon;
using ASP.ConfigCommon;
using ASP.Models.Admin.Accounts;
using ASP.Models.Admin.Auths;
using ASP.Models.ASPModel;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Threading.Tasks;
using System.Transactions;

namespace ASP.Controllers.Admin
{
    public class AuthController : Controller, AuthListenerInterface
    {
        private readonly ASPDbContext _context;
        private UserManager<ApplicationUser> _userManager;
        private SignInManager<ApplicationUser> _signInManager;
        private AuthRepositoryInterface auth;
        // GET: AuthController
        public AuthController(ASPDbContext context, UserManager<ApplicationUser> userMgr, SignInManager<ApplicationUser> signinMgr, AuthRepositoryInterface auth)
        {
            _context = context;
            _userManager = userMgr;
            _signInManager = signinMgr;
            this.auth = auth;
        }

        [HttpGet]
        [Route("admin")]
        [Route("Login")]
        [Route("Auth")]
        [Route("Auth/Index")]
        [Route("Auth/Login", Name = "admin.auths.login")]
        [AllowAnonymous]
        public IActionResult Login(string returnUrl)
        {
            Login login = new Login();
            login.ReturnUrl = returnUrl;
            //
            var chkLogged = _userManager.FindByNameAsync(_userManager.GetUserName(User));//.Result.UserName
            if (chkLogged.Exception == null)
            {
                if (login.ReturnUrl != null)
                {
                    return Redirect(login.ReturnUrl);
                }
                else
                {
                    return RedirectToAction("Index", "Dashboard");
                }
            }
            return View("../Admin/Auth/Login", login);
        }
        
        [HttpPost]
        [ValidateAntiForgeryToken]
        [Route("Login", Name = "admin.auth.login")]
        [AllowAnonymous]
        public async Task<IActionResult> Login(Login login)
        {
            //var serviceEz = new ServiceReferencePPS.ServiceManPowerEzClient();
            //var serviceEz = new ServiceEz.ServiceManPowerEzClient();
            var serviceEzV4 = new ServiceEzV4.Service1Client();
            //bool chkAccAD = serviceEz.CheckAccountAdAsync(login.UserName, login.PassWord).Result;
            /**
             * API APAD
             * **/
            /**
             * V4
             * **/
            //authenticates against your local machine - for development time

            bool chkAccAD = false;
            try
            {
                //
                string apiUrl = "http://10.73.131.59:1113/api/login/LoginByAzAd";
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(apiUrl);
                request.Method = "POST";
                request.ContentType = "application/json";
                request.Accept = "application/json";
                string postData = JsonConvert.SerializeObject(new
                {
                    employeeCode = login.UserName,
                    password = login.PassWord
                });

                // Ghi dữ liệu POST vào request stream
                using (var streamWriter = new StreamWriter(request.GetRequestStream()))
                {
                    streamWriter.Write(postData);
                    streamWriter.Flush();
                    streamWriter.Close();
                }
                using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                {
                    using (StreamReader reader = new StreamReader(response.GetResponseStream()))
                    {
                        string responseData = reader.ReadToEnd();
                        // Xử lý dữ liệu từ API ở đây
                        ResponseAuthenModel responseAuthen = JsonConvert.DeserializeObject<ResponseAuthenModel>(responseData);
                        if (responseAuthen.result)
                        {
                            chkAccAD = true;
                        }
                    }
                }
                //return false;
            }
            catch (Exception)
            {
                chkAccAD = false;
            }
            //
            if (chkAccAD)
            {
                #region login with AD
                login.UserName = login.UserName.ToLower().Replace("dmvn", "").Replace("vn", "");
                var infoEmp = serviceEzV4.GetEmployeeInfoV4Async(login.UserName, false).Result;
                /* ** 
                 * co tai khoan may tinh OK
                 * ktra xem co username tren he thong chua?
                 * chua co => add new
                 * da ton tai => login
                 */
                string strHome = "/";
                ApplicationUser appUser = await _userManager.FindByNameAsync(login.UserName);
                //
                if (appUser != null)
                {
                    // ktra username thuoc role nao?
                    var chkUserRole = _context.UserRoles.FirstOrDefault(f => f.UserId == appUser.Id);
                    if (chkUserRole != null)
                    {
                        var chkRoleDefault = _context.Roles.Any(f => f.Id == chkUserRole.RoleId && f.DefaultRole == true);
                        if (!chkRoleDefault) // Role = User
                        {
                            strHome = "/admin/dashboard";
                        }
                    }
                    /**
                     * da co username tren he thong: ktra mat khau
                     * neu sai pass => cap nhat password & login
                     * **/
                    await _signInManager.SignOutAsync();
                    var signinResult = await _signInManager.PasswordSignInAsync(appUser, login.PassWord, false, lockoutOnFailure: false);//login.PassWord
                    if (signinResult.Succeeded)
                    {
                        if (string.IsNullOrEmpty(login.ReturnUrl))
                        {
                            return RedirectToAction("Index", "Home");
                        }
                        else
                        {
                            return Redirect(login.ReturnUrl);
                        }
                    }
                    else
                    {
                        #region cap nhat password & login
                        if (!string.IsNullOrEmpty(login.PassWord))
                        {
                            var password = new PasswordHasher<ApplicationUser>();
                            var hashed = password.HashPassword(appUser, login.PassWord);
                            appUser.PasswordHash = hashed;
                            appUser.Status = (int)EnumStatusUser.Active;
                            _context.SaveChanges();
                            // login
                            // da co username tren he thong: ktra mat khau
                            await _signInManager.SignOutAsync();
                            var signinResult2 = await _signInManager.PasswordSignInAsync(appUser, login.PassWord, false, lockoutOnFailure: false);//login.PassWord
                            if (signinResult2.Succeeded)
                            {
                                //return Redirect(login.ReturnUrl ?? strHome);
                                if (string.IsNullOrEmpty(login.ReturnUrl))
                                {
                                    return RedirectToAction("Index", "Home");
                                }
                                else
                                {
                                    return Redirect(login.ReturnUrl);
                                }
                            }
                            else
                            {
                                // ko cap nhat dc pass => dang nhap failed
                                ModelState.AddModelError(nameof(login.UserName), "Lỗi: Tài khoản hoặc Mật khẩu không hợp lệ.");
                            }
                            //
                        }
                        else
                        {
                            // ko co password
                            ModelState.AddModelError(nameof(login.UserName), "Lỗi: Tài khoản hoặc Mật khẩu không hợp lệ.");
                        }
                        #endregion
                    }
                    //
                }
                else
                {
                    /**
                     * chua co username tren he thong
                     * them moi tai khoan & login
                     * 
                     * **/
                    Register rgtObj = new Register()
                    {
                        UserName = login.UserName,
                        PassWord = login.PassWord,
                        FullName = infoEmp != null ? infoEmp.FullName : "",
                    };
                    var result = await auth.CheckAccountAD(rgtObj);
                    if (result)
                    {
                        /**
                         * login
                         * da co username tren he thong: ktra mat khau
                         * **/
                        ApplicationUser appUser2 = await _userManager.FindByNameAsync(login.UserName);
                        await _signInManager.SignOutAsync();
                        var signinResult2 = await _signInManager.PasswordSignInAsync(appUser2, login.PassWord, false, lockoutOnFailure: false);//login.PassWord
                        if (signinResult2.Succeeded)
                        {
                            //return Redirect(login.ReturnUrl ?? strHome);
                            if (string.IsNullOrEmpty(login.ReturnUrl))
                            {
                                return RedirectToAction("Index", "Home");
                            }
                            else
                            {
                                return Redirect(login.ReturnUrl);
                            }
                        }
                        else
                        {
                            // ko cap nhat dc pass => dang nhap failed
                            ModelState.AddModelError(nameof(login.UserName), "Lỗi: Tài khoản hoặc Mật khẩu không hợp lệ.");
                        }
                    }
                    else
                    {
                        // ko them dc tai khoan => dang nhap failed
                        ModelState.AddModelError(nameof(login.UserName), "Lỗi: Tài khoản hoặc Mật khẩu không hợp lệ.");
                    }
                }
                #endregion
                //
            }
            else
            {
                #region system
                //
                var chkUser = _context.Users.FirstOrDefault(f => f.UserName == login.UserName);

                #region login
                if (chkUser != null)
                {
                    if (chkUser.Status != (int)EnumStatusUser.Active)
                    {
                        ModelState.AddModelError(nameof(login.UserName), "Lỗi: Tài khoản chưa được kích hoạt bởi quản trị viên.");
                    }
                    #region da co tai khoan tren he thong
                    if (ModelState.IsValid)
                    {
                        ApplicationUser appUser = await _userManager.FindByNameAsync(login.UserName);
                        if (appUser != null)
                        {
                            // ktra username thuoc role nao?
                            string strHome = "/";
                            var chkUserRole = _context.UserRoles.FirstOrDefault(f => f.UserId == appUser.Id);
                            if (chkUserRole != null)
                            {
                                var chkRoleDefault = _context.Roles.Any(f => f.Id == chkUserRole.RoleId && f.DefaultRole == true);
                                if (!chkRoleDefault) // Role = User
                                {
                                    strHome = "/admin/dashboard";
                                }
                            }
                            //
                            await _signInManager.SignOutAsync();
                            var signinResult = await _signInManager.PasswordSignInAsync(appUser, login.PassWord, false, lockoutOnFailure: false);//login.PassWord
                            if (signinResult.Succeeded)
                            {
                                //return Redirect(login.ReturnUrl ?? strHome);
                                if (string.IsNullOrEmpty(login.ReturnUrl))
                                {
                                    return RedirectToAction("Index", "Home");
                                }
                                else
                                {
                                    return Redirect(login.ReturnUrl);
                                }
                            }
                            //
                        }
                        ModelState.AddModelError(nameof(login.UserName), "Lỗi: Tài khoản hoặc Mật khẩu không hợp lệ.");
                    }
                    #endregion
                }
                else
                {
                    // user chua ton tai
                    ModelState.AddModelError(nameof(login.UserName), "Lỗi: Tài khoản hoặc Mật khẩu không hợp lệ.");
                    return View("../Admin/Auth/Login", login);
                }
                #endregion
                //
                #endregion
            }
            return View("../Admin/Auth/Login", login);
        }
        [HttpGet]
        [Route("Logout", Name = "admin.auth.logout")]
        [Route("Auth/Logout")]
        public async Task<IActionResult> Logout()
        {
            await _signInManager.SignOutAsync();
            return RedirectToAction("Login", "Auth");
        }
        [HttpGet]
        [Route("Register")]
        [Route("Auth/Register", Name = "admin.auths.register")]
        public IActionResult Register()
        {
            return View("../Admin/Auth/Register");
        }
        [HttpPost]
        [ValidateAntiForgeryToken]
        [Route("Register", Name = "admin.auths.store")]
        public async Task<IActionResult> Store(Register register)
        {

            var chkUser = _context.Users.FirstOrDefault(f => f.UserName == register.UserName);
            #region validation
            if (chkUser != null)
            {
                ModelState.AddModelError(nameof(register.UserName), "Tài khoản đã tồn tại trên hệ thống.");
            }
            if (string.IsNullOrEmpty(register.PassWord) || string.IsNullOrEmpty(register.RePassWord))
            {
                ModelState.AddModelError("Password", "Mật khẩu không để trống.");
            }
            ModelState.Remove("Avatar");
            if (!ModelState.IsValid)
            {
                return RegisterAccountFails(register);
            }
            #endregion
            #region add new
            await auth.CreateAccount(this, register);
            #endregion
            return View("../Admin/Auth/Register");
        }
        [HttpPost]
        public IActionResult RegisterAccountSuccess()
        {
            TempData["mess-type"] = "success";
            TempData["mess-detail"] = BaseController.BaseMessage("create_success");
            return RedirectToAction(nameof(Login));
        }
        [HttpPost]
        public IActionResult RegisterAccountFails(Register user)
        {
            TempData["mess-type"] = "error";
            TempData["mess-detail"] = BaseController.BaseMessage("create_fails");
            return View("../Admin/Auth/Register", user);
        }
   
    }

}
