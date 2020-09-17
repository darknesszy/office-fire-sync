# Office FireSync
Sychronise FireStore with Office suite documents.

# Setup
1. Add Environmental variables:
	* GOOGLE_APPLICATION_CREDENTIALS: Json file downloaded from Firebase permissions panel.
	* PROJECT_ID: Project ID from project settings panel.
	* MEDIA_PATH: Folder with images that are used with the data (Not ready, just add a random path for now).
2. Restore dependencies with dotnet restore
3. Run project with CLI command:
	* Shopify: -r "<path to the xlsx file>" -e "shopify"
