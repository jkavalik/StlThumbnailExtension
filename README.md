# StlThumbnailExtension

This library is a Windows Explorer Shell Extension to show thumbnails of STL files (3D models) by rendering a preview of the model geometry.

## Installation

Download the release, extract contents, and run `register.bat` as Administrator.
Keep the DLL in the same folder unless you unregister it.

## Features

- Shows rendered 3D previews for `.stl` files (ASCII and Binary)
- Integrates with Windows Explorer as a shell extension

## Limitations

- Thumbnail is a simple shaded projection (not full 3D, no colors)
- Large or complex models may take a few moments to render
- Only first 500k triangles are rendered for performance

## Uninstallation

Run `unregister.bat` as Administrator.