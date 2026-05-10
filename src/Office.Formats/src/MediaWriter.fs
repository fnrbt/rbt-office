module MediaWriter

open Fable.Core
open Fable.Core.JsInterop
open XmlGenerator

/// Image format
type ImageFormat =
    | PNG
    | JPEG
    | GIF
    | BMP

/// Image data
type ImageData = {
    Data: byte[]  // Image binary data
    Format: ImageFormat
    Width: int option  // in EMUs (English Metric Units)
    Height: int option  // in EMUs
}

/// Get image extension
let getImageExtension (format: ImageFormat) : string =
    match format with
    | PNG -> "png"
    | JPEG -> "jpeg"
    | GIF -> "gif"
    | BMP -> "bmp"

/// Get image content type
let getImageContentType (format: ImageFormat) : string =
    match format with
    | PNG -> "image/png"
    | JPEG -> "image/jpeg"
    | GIF -> "image/gif"
    | BMP -> "image/bmp"

/// Convert pixels to EMUs (1 pixel = 9525 EMUs at 96 DPI)
let pixelsToEmus (pixels: int) : int =
    pixels * 9525

/// Generate image relationship
let generateImageRelationship (imageId: int) (fileName: string) : string =
    selfClosing "Relationship" [
        ("Id", sprintf "rId%d" imageId)
        ("Type", RelationshipTypes.image)
        ("Target", sprintf "media/%s" fileName)
    ]

/// Generate drawing XML for inline image
let generateInlineImage (imageId: int) (width: int) (height: int) (imageName: string) : string =
    // Simplified inline image (full implementation would be more complex)
    let drawing = element "w:drawing" [] [
        element "wp:inline" [
            ("xmlns:wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing")
        ] [
            element "wp:extent" [("cx", string width); ("cy", string height)] []
            element "wp:docPr" [("id", string imageId); ("name", imageName)] []
            element "a:graphic" [("xmlns:a", "http://schemas.openxmlformats.org/drawingml/2006/main")] [
                element "a:graphicData" [("uri", "http://schemas.openxmlformats.org/drawingml/2006/picture")] [
                    element "pic:pic" [("xmlns:pic", "http://schemas.openxmlformats.org/drawingml/2006/picture")] [
                        element "pic:blipFill" [] [
                            selfClosing "a:blip" [("r:embed", sprintf "rId%d" imageId)]
                        ]
                        element "pic:spPr" [] [
                            element "a:xfrm" [] [
                                selfClosing "a:off" [("x", "0"); ("y", "0")]
                                selfClosing "a:ext" [("cx", string width); ("cy", string height)]
                            ]
                            element "a:prstGeom" [("prst", "rect")] []
                        ]
                    ]
                ]
            ]
        ]
    ]
    drawing

/// Generate run with image
let generateImageRun (imageId: int) (width: int) (height: int) (imageName: string) : string =
    element "w:r" [] [
        generateInlineImage imageId width height imageName
    ]

/// Create default image dimensions (4 inches wide)
let defaultImageWidth = pixelsToEmus 384  // 4 inches at 96 DPI
let defaultImageHeight = pixelsToEmus 288  // 3 inches at 96 DPI

/// Helper to create image data from base64
[<Emit("Buffer.from($0, 'base64')")>]
let base64ToBytes (base64: string) : byte[] = jsNative

/// Helper to create image data (for Node.js environment)
let createImageData (base64Data: string) (format: ImageFormat) (widthPixels: int option) (heightPixels: int option) : ImageData =
    {
        Data = base64ToBytes base64Data
        Format = format
        Width = widthPixels |> Option.map pixelsToEmus
        Height = heightPixels |> Option.map pixelsToEmus
    }

/// Note: Full image embedding support requires updating DocumentWriter
/// to handle images in paragraphs and manage media relationships.
/// This module provides the infrastructure for Phase 4 implementation.
