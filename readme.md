# Onedrive-DB

## What is it?
This is a small library that uses the graph api to read and write key-value based data to OneDrive.

## How does it work?
It stores json-files under the Onedrive-App folder.
These files can be read and written to to store any kind of data (as long as it can be parse to JSON)

## Getting started
Requirements:
* Make sure to include [hello.js](https://adodson.com/hello.js/) and this library (in this order) + the redirect.html
* Create an app in the [Azure portal](https://portal.azure.com)
* Give this app "User.Read" and "Files.ReadWrite.AppFolder" permissions
* Set the redirect-url to the desired url
* Copy the App-ID

Create a new db connection (make sure to fill in the id and name).
```
const db = new OnedriveDB('<your-app-id-here>', '<your-app-name-here>');
```
Next, initalize the connection.
```
db.init().then(data => {});
```
Now, you can get the data at any time by using db.data.

You can reload the data like this:
```
db.load().then(data => {});
```

Once you made changes to the data, save the result:
```
db.save();
```

