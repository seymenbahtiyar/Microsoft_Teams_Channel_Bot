# Microsoft Teams Channel Bot


This is a simple bot written in C# that uses the Microsoft Graph API to respond to messages in a Microsoft Teams channel.

## Dependencies

- Microsoft.Identity.Client 4.50.0
- RestSharp 108.0.3
- Newtonsoft.Json 13.0.2

## Configuration

Before running the program, you need to set the following variables with your own values:

- `tenantID`: The ID of your Azure AD tenant.
- `clientID`: The ID of your Azure AD application.
- `userEmail`: The email address of the user you want to authenticate as.
- `userPassword`: The password of the user you want to authenticate as.
- `teamID`: The ID of the Microsoft Teams team you want to interact with.
- `channelID`: The ID of the channel within the team you want to interact with.

## Example

<img width="304" alt="image" src="https://user-images.githubusercontent.com/110940406/225104148-23ea9122-b75b-4ef6-9d47-4558fb5f357e.png">


## How it works

The program enters an infinite loop where it checks for new messages in the specified channel every 5 seconds. If it finds a new message, it calls the `RespondToMessage` method which sends a reply to that message.

The reply is hardcoded as "Reply from C#!" but you can change it to whatever you want.



## Running the program

To run the program, simply build and run it using your preferred C# development environment.

## License

[MIT](https://github.com/seymenbahtiyar/Microsoft_Teams_Channel_Bot/blob/main/LICENSE)
