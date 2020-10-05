# Office FireSync
Sychronise FireStore with Open XML documents.

** Currently only support manual run. Future updates will add change detection, allowing it to be hosted on a server.


# Environment Variables
>GOOGLE_APPLICATION_CREDENTIALS

File path to the API key json file downloaded from Firebase project settings.

>PROJECT_ID

Project ID found in Firebase project settings.

>MEDIA_PATH

Folder with images that are used with the data (Not ready, just add a random path for now).

# Quick Start

Restore project dependencies:

`dotnet restore`

Create a file named .env in the root directory and fill in the environment variables. Run the tool.

`dotnet run -r "C:\Users\username\Documents\xxx.xlsx" -s "product" -e "shopify"`

# Usage
`dotnet run [COMMAND] [ARGS...]`

The `--excel` command parses through the provided excel document converting it to add to or update a FireStore collection.

`--excel` currently supports `table` based, `sheet` based, and `shopify` cvs based conversion with each respective term supplied as option to the command.

The `--word` command parses through the provided word document to extract text and style information to add to or update a FireStore collection.

`--word` currently supports `html`, converting part of the word document into HTML with styling and `heading` creating Firebase arrays from sections separated by heading.

# Arguments
<table>
    <thead>
        <tr>
            <th>Name, shorthand</th>
            <th>Description</th>
        </tr>
    </thead>
    <tbody>
		<tr>
            <td>
                <code>--read, -r</code>
            </td>
            <td>
                Specifies the file path to the document that will be parsed. Required.
            </td>
        </tr>
        <tr>
            <td>
                <code>--sync, -s</code>
            </td>
            <td>
                Specifies the Collection name to be synced to Stripe. Required.
            </td>
        </tr>
    </tbody>
</table>