namespace Inviter
{
    using System;
    using System.IO;
    using System.Threading.Tasks;
    using System.Net.Http;
    using System.Net.Http.Headers;

    public class Program
    {
        private static readonly HttpClient client = new HttpClient();
        private const string funcURL = "https://{app-name-here}.azurewebsites.net/api/{fuction-name-here}?email={0}&name={1}&code={2}&invitation=true";
        private const string funcCode = "{azure-function-code-here}";

        private static async Task ProcessInvitation(string name, string email)
        {
            client.DefaultRequestHeaders.Accept.Clear();
            client.DefaultRequestHeaders.Accept.Add(
                new MediaTypeWithQualityHeaderValue("application/json"));
            client.DefaultRequestHeaders.Add("User-Agent", "Invitation Client");


            var nameUrl = System.Net.WebUtility.UrlEncode(name);
            var emailUrl = System.Net.WebUtility.UrlEncode(email);
            Console.WriteLine("Name: " + nameUrl);
            Console.WriteLine("email: " + emailUrl);
            await client.GetAsync(string.Format(funcURL ,emailUrl, nameUrl, funcCode));
        }

        static void Main(string[] args)
        {
            Console.WriteLine("Processing");

            using(var reader = new StreamReader(args[0]))
            {
                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    var values = line.Split(';');

                    ProcessInvitation(values[0], values[1]).Wait();
                    Console.WriteLine("====");
                }
            }
        }
    }
}
