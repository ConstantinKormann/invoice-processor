# Icon Assets

This directory can contain custom application icons for the InvoiceProcessor app.

## Generating Icons

For Windows (.ico):
- Create a 256x256 PNG image with your icon design
- Use an online converter or tool like ImageMagick to convert to .ico format
- Name it `icon.ico` and place it in this directory

For macOS (.icns):
- Create a 1024x1024 PNG image with your icon design
- Use Icon Composer or `iconutil` to create the .icns file
- Name it `icon.icns` and place it in this directory

The build scripts will automatically use these icons if they're present.
