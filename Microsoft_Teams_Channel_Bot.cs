using Microsoft.Identity.Client; // Import the Microsoft Identity Client library
using RestSharp; // Import the RestSharp library
using System; // Import the System library
using System.Threading.Tasks; // Import the System.Threading.Tasks library

namespace TeamsBot // Define the namespace
{
    public class Program11 // Define the class
    {
        static string tenantID = "tenantID"; // Declare a static string variable for the tenant ID
        static string clientID = "clientID"; // Declare a static string variable for the client ID
        static string userEmail = "userEmail"; // Declare a static string variable for the user email
        static string userPassword = "userPassword"; // Declare a static string variable for the user password
        static string teamID = "teamID"; // Declare a static string variable for the team ID
        static string channelID = "channelID"; // Declare a static string variable for the channel ID
        static string microsoftGraphAPIBaseURL = "https://graph.microsoft.com/v1.0/"; // Declare a static string variable for the Microsoft Graph API base URL
        static AuthenticationResult result; // Declare a static variable for the authentication result

        public static async Task Main(string[] args) // Define the main method
        {
            string authority = $"https://login.microsoftonline.com/{tenantID}"; // Define a string variable for the authority URL
            var app = PublicClientApplicationBuilder
                .Create(clientID)
                .WithAuthority(authority)
                .Build(); // Create a public client application with the client ID and authority URL
            result = await app.AcquireTokenByUsernamePassword(
                    new[] { "User.Read", "ChannelMessage.Read.All" },
                    userEmail,
                    userPassword
                )
                .ExecuteAsync(); // Acquire an access token by using username and password authentication and assign it to the result variable
            var client = new RestClient(microsoftGraphAPIBaseURL); // Create a new RestClient object with the Microsoft Graph API base URL
            string lastMessageId = null; // Declare a string variable for the last message ID and initialize it to null
            while (true) // Start an infinite loop
            {
                var request = new RestRequest(
                    $"teams/{teamID}/channels/{channelID}/messages",
                    Method.Get
                ); // Create a new RestRequest object with the endpoint to get messages from a channel and set the method to GET
                request.AddHeader("Authorization", $"Bearer {result.AccessToken}"); // Add an authorization header with the access token to the request
                var response = await client.ExecuteAsync(request); // Execute the request asynchronously and assign the response to a variable
                if (response.IsSuccessful) // Check if the response is successful
                {
                    dynamic messages = Newtonsoft.Json.JsonConvert.DeserializeObject(
                        response.Content
                    ); // Deserialize the response content as a dynamic object using Newtonsoft.Json library
                    if (messages.value.Count > 0) // Check if there are any messages in the response
                    {
                        string messageId = messages.value[0].id; // Assign the ID of the first message to a variable
                        if (messageId != lastMessageId && lastMessageId != null) // Check if the message ID is different from the last message ID and if the last message ID is not null
                        {
                            RespondToMessage(messages.value[0]); // Call the RespondToMessage method with the first message as an argument
                        }
                        lastMessageId = messageId; // Update the last message ID with the current message ID
                    }
                }
                await Task.Delay(5000); // Wait for 5 seconds before repeating the loop
            }
        }

        private static async void RespondToMessage(dynamic message) // Define a private static method that takes a dynamic object as an argument and responds to it asynchronously
        {
            var client = new RestClient(microsoftGraphAPIBaseURL); // Create a new RestClient object with the Microsoft Graph API base URL
            var request = new RestRequest(
                $"teams/{teamID}/channels/{channelID}/messages/{message.id}/replies",
                Method.Post
            ); // Create a new RestRequest object with the endpoint to post a reply to a message and set the method to POST
            request.AddHeader("Authorization", $"Bearer {result.AccessToken}"); // Add an authorization header with the access token to the request
            request.AddHeader("Content-Type", "application/json"); // Add a content-type header with application/json value to the request
            var body = new { body = new { content = "Reply from C#!" } }; // Define an anonymous object for the request body with a content property that contains the reply text
            request.AddJsonBody(body); // Add the body object as JSON to the request
            var response = await client.ExecuteAsync(request); // Execute the request asynchronously and assign the response to a variable

            if (response.IsSuccessful) // Check if the response is successful
            {
                Console.WriteLine("Reply sent successfully"); // Write a message to the console indicating success
            }
            else // If the response is not successful
            {
                Console.WriteLine("Failed to send reply"); // Write a message to the console indicating failure
            }
        }
    }
}
