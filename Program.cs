using System;
using System.Timers;
using RestSharp;
using Newtonsoft.Json.Linq;
using System.Windows.Forms;
using System.Drawing;
using System.Diagnostics;
using OpenXmlPowerTools;

class Program
{
    private static System.Timers.Timer pollTimer;
    private static string linkedInClientId = "";
    private static string linkedInClientSecret = "";
    private static string linkedInAccessToken = "";
    private static string redirectUri = "";
    private static string linkedInApiUrl = "";
    private static string memberUrn = "";

    static void Main(string[] args)
    {

        StartOAuthFlow();


        pollTimer = new System.Timers.Timer(60000);
        pollTimer.Elapsed += PollLinkedIn;
        pollTimer.AutoReset = true;
        pollTimer.Enabled = true;

        Console.WriteLine("Started LinkedIn Notification App. Press Enter to stop.");
        Console.ReadLine();
    }

    private static void StartOAuthFlow()
    {

        var authorizationUrl = $"https://www.linkedin.com/oauth/v2/authorization?response_type=code&client_id={Uri.EscapeDataString(linkedInClientId)}&redirect_uri={Uri.EscapeDataString(redirectUri)}&scope=openid%20profile%20email%20w_member_social";

        Console.WriteLine("Opening LinkedIn authorization URL...");
        Console.WriteLine(authorizationUrl);


        Process.Start(new ProcessStartInfo("cmd", $"/c start {authorizationUrl}") { CreateNoWindow = true });


        Console.WriteLine("Please enter the authorization code from the URL after login:");
        string authorizationCode = Console.ReadLine();


        GetAccessToken(authorizationCode);
    }


    private static void GetAccessToken(string authorizationCode)
    {
        var client = new RestClient("https://www.linkedin.com/oauth/v2/accessToken");
        var request = new RestRequest();
        request.Method = Method.Post;

        request.AddParameter("grant_type", "authorization_code");
        request.AddParameter("code", authorizationCode);
        request.AddParameter("redirect_uri", redirectUri);
        request.AddParameter("client_id", linkedInClientId);
        request.AddParameter("client_secret", linkedInClientSecret);

        var response = client.Execute(request);
        var content = JObject.Parse(response.Content);

        if (response.IsSuccessful)
        {
            linkedInAccessToken = content["access_token"]?.ToString();
            Console.WriteLine("Access token received successfully.");


            GetUserInfo(linkedInAccessToken);
        }
        else
        {
            Console.WriteLine("Failed to obtain access token.");
            Console.WriteLine(response.Content);
        }
    }

    private static void GetUserInfo(string accessToken)
    {
        var client = new RestClient("https://api.linkedin.com/v2/userinfo");
        var request = new RestRequest();
        request.Method = Method.Get;
        request.AddHeader("Authorization", $"Bearer {accessToken}");

        var response = client.Execute(request);
        if (response.IsSuccessful)
        {
            var userInfo = JObject.Parse(response.Content);
            var id = userInfo["sub"]?.ToString();
            memberUrn = $"urn:li:person:{id}";
            Console.WriteLine($"Member URN: {memberUrn}");


            PostOnLinkedIn(accessToken, memberUrn);
        }
        else
        {
            Console.WriteLine($"Failed to retrieve user info: {response.StatusDescription}");
            Console.WriteLine(response.Content);
        }
    }


    private static void PostOnLinkedIn(string accessToken, string memberUrn)
    {
        var client = new RestClient("https://api.linkedin.com/v2/ugcPosts");
        var request = new RestRequest();
        request.Method = Method.Post;


        request.AddHeader("Authorization", $"Bearer {accessToken}");
        request.AddHeader("X-Restli-Protocol-Version", "2.0.0");


        var postBody = new
        {
            author = memberUrn,
            lifecycleState = "PUBLISHED",
            specificContent = new Dictionary<string, object>
            {
                ["com.linkedin.ugc.ShareContent"] = new
                {
                    shareCommentary = new
                    {
                        text = "This is a test post from my LinkedIn Notifier App by Prof.C Murwamuila!"
                    },
                    shareMediaCategory = "NONE"
                }
            },
            visibility = new Dictionary<string, string>
            {
                ["com.linkedin.ugc.MemberNetworkVisibility"] = "PUBLIC"
            }
        };

        request.AddJsonBody(postBody);

        var response = client.Execute(request);

        if (response.IsSuccessful)
        {
            Console.WriteLine("Post created successfully.");
        }
        else
        {
            Console.WriteLine($"Failed to create post: {response.StatusCode}");
            Console.WriteLine(response.Content);
        }
    }


    private static void PostArticleOnLinkedIn(string accessToken, string memberUrn, string articleUrl, string articleTitle, string articleDescription)
    {
        var client = new RestClient("https://api.linkedin.com/v2/ugcPosts");
        var request = new RestRequest();
        request.Method = Method.Post;


        request.AddHeader("Authorization", $"Bearer {accessToken}");
        request.AddHeader("X-Restli-Protocol-Version", "2.0.0");


        var postBody = new
        {
            author = memberUrn,
            lifecycleState = "PUBLISHED",
            specificContent = new
            {
                ShareContent = new
                {
                    shareCommentary = new
                    {
                        text = "Learning more about LinkedIn by reading the LinkedIn Blog!"
                    },
                    shareMediaCategory = "ARTICLE",
                    media = new[]
                    {
                        new
                        {
                            status = "READY",
                            description = new { text = articleDescription },
                            originalUrl = articleUrl,
                            title = new { text = articleTitle }
                        }
                    }
                }
            },
            visibility = new
            {
                MemberNetworkVisibility = "PUBLIC"
            }
        };

        request.AddJsonBody(postBody);

        var response = client.Execute(request);

        if (response.IsSuccessful)
        {
            Console.WriteLine("Article post created successfully.");
        }
        else
        {
            Console.WriteLine($"Failed to create article post: {response.StatusCode}");
            Console.WriteLine(response.Content);
        }
    }

    private static void PollLinkedIn(object sender, ElapsedEventArgs e)
    {
        if (string.IsNullOrEmpty(linkedInAccessToken))
        {
            Console.WriteLine("Access token not found. Please authenticate.");
            return;
        }

        Console.WriteLine("Polling LinkedIn for updates...");
    }

    private static void ShowNotification(string title, string message)
    {

        NotifyIcon notifyIcon = new NotifyIcon();
        notifyIcon.Icon = SystemIcons.Information;
        notifyIcon.Visible = true;
        notifyIcon.BalloonTipTitle = title;
        notifyIcon.BalloonTipText = message;
        notifyIcon.ShowBalloonTip(3000);
    }
}
