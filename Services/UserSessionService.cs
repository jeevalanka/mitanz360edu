using MITANZ360Edu.Web.Models;

namespace MITANZ360Edu.Web.Services;

public class UserSessionService
{
    public UserSessionModel CurrentUser { get; private set; } = new();

    public bool IsLoggedIn =>
        CurrentUser.IsAuthenticated;

    public void SetUser(UserSessionModel user)
    {
        CurrentUser = user;
    }

    public void Clear()
    {
        CurrentUser = new UserSessionModel();
    }
}