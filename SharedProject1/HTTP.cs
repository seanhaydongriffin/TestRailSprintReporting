using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections;
using System.IO;
using System.Net;
using System.Web;
using System.Collections.Specialized;
using System.Security.Cryptography;
using Newtonsoft.Json.Linq;
using System.Globalization;
using System.Reflection;
//using System.Net.Http;
using System.Net.Configuration;

namespace SharedProject
{

    public static class HTTP
    {


        // ===============================================================================
        // Name...........:	UrlEncode()
        // Description....:	Encodes a Url (or other similar data, for instance HTTP).
        // Syntax.........:	UrlEncode(this String url, int method)
        // Parameters.....:	method		- Optional: A number denoting the method to use to encode the Url.
        //								    1 = via EscapeUriString (default)
        //                                  2 = via EscapeDataString
        //                                  3 = via UrlEncode
        //                                  4 = via UrlPathEncode
        //                                  5 = via HtmlEncode
        //                                  6 = via HtmlAttributeEncode
        // Return values..: The encoded url.
        // Remarks........:	None.
        // ==========================================================================================

        public static String UrlEncode(this String url, int method = 1)
        {
            switch (method)
            {
                case 1:

                    return Uri.EscapeUriString(url);

                case 2:

                    return Uri.EscapeDataString(url);

                case 3:

                    return HttpUtility.UrlEncode(url);

                case 4:

                    return HttpUtility.UrlPathEncode(url);

                case 5:

                    return HttpUtility.HtmlEncode(url);

                case 6:

                    return HttpUtility.HtmlAttributeEncode(url);

                case 7:

                    return System.Net.WebUtility.UrlEncode(url);
            }

            return null;

        }



    }
}
