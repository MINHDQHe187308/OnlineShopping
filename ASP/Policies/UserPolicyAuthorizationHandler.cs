using ASP.ConfigCommon;
using ASP.Models;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Authorization.Infrastructure;
using Microsoft.AspNetCore.Identity;
using Newtonsoft.Json;
using ASP.BaseCommon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ASP.Policies
{
    public class UserPolicyAuthorizationHandler :
    AuthorizationHandler<OperationAuthorizationRequirement, DocumentAuth>
    {
        protected override Task HandleRequirementAsync(AuthorizationHandlerContext context,
                                                       OperationAuthorizationRequirement requirement,
                                                       DocumentAuth resource)
        {
            var getUserClaim = context.User.Claims.FirstOrDefault(c => c.Type == requirement.Name);
            if (getUserClaim != null)
            {
                if (getUserClaim.Value == "1")
                {
                    context.Succeed(requirement);
                }
            }
            return Task.CompletedTask;
        }
    }
}
