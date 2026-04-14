# Deploy Guide

## 1. Prepare Google Sheets

Create one spreadsheet with these sheet names:

- `products`
- `stock`
- `orders`
- `order_items`

The script will create the headers automatically if the sheets are empty.

## 2. Prepare Google Drive

Create one folder for product images and copy its folder ID from the URL.

## 3. Deploy Google Apps Script

1. Open [script.google.com](https://script.google.com).
2. Create a new project.
3. Replace the default file with the contents of [Code.gs](/var/www/html/quanlyaodai/Code.gs).
4. Set:
   - `SPREADSHEET_ID`
   - `DRIVE_FOLDER_ID`
5. Save the project.
6. Click `Deploy` -> `New deployment`.
7. Choose `Web app`.
8. Set:
   - Execute as: `Me`
   - Who has access: `Anyone`
9. Deploy and copy the Web App URL.

## 4. Connect the frontend

Open [app.js](/var/www/html/quanlyaodai/app.js) and replace:

```js
https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec
```

with your deployed Web App URL.

## 5. Publish to GitHub Pages

1. Push this repo to GitHub.
2. In the GitHub repo, open `Settings` -> `Pages`.
3. Set source to your main branch root.
4. Save and wait for the Pages URL.

## 6. Install on iPhone

1. Open the GitHub Pages URL in Safari.
2. Tap `Share`.
3. Tap `Add to Home Screen`.
4. Launch the app from the Home Screen.

## 7. Expected API actions

The frontend calls these Apps Script actions:

- `addProductFull`
- `getProducts`
- `createOrder`
- `stockIn`
