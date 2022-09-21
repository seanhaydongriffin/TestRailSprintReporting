

using System.Net;
using System.IO;
using System.Text;

namespace SharedProject.TestRail
{
    public class WebClient
    {
        private string m_user;
        private string m_password;
        private string m_url;
        private CookieContainer cookieContainer = new CookieContainer();

        public WebClient(string base_url)
        {
            if (!base_url.EndsWith("/"))
            {
                base_url += "/";
            }

            this.m_url = base_url + "index.php?/";
        }

        /**
         * Get/Set User
         *
         * Returns/sets the user used for authenticating the API requests.
         */
        public string User
        {
            get { return this.m_user; }
            set { this.m_user = value; }
        }

        /**
         * Get/Set Password
         *
         * Returns/sets the password used for authenticating the API requests.
         */
        public string Password
        {
            get { return this.m_password; }
            set { this.m_password = value; }
        }

        public void Login()
        {
            var request = (HttpWebRequest)WebRequest.Create(@"https://janison.testrail.com/index.php?/auth/login/");
            var postData = "name=" + this.m_user + "&password=" + this.m_password + "&rememberme=1";
            var data = Encoding.ASCII.GetBytes(postData);
            request.Method = "POST";
            request.ContentType = "application/x-www-form-urlencoded";
            request.ContentLength = data.Length;
            request.ServicePoint.Expect100Continue = false;
            request.CookieContainer = cookieContainer;

            using (var stream = request.GetRequestStream())
            {
                stream.Write(data, 0, data.Length);
            }

            var response = (HttpWebResponse)request.GetResponse();
        }

        public string SendGet(string uri)
        {
            string url = this.m_url + uri;
            var request = (HttpWebRequest)WebRequest.Create(url);
            request.Method = "GET";
            request.CookieContainer = cookieContainer;
            var response = (HttpWebResponse)request.GetResponse();
            var responseString = new StreamReader(response.GetResponseStream()).ReadToEnd();
            return responseString;
        }


    }
}
