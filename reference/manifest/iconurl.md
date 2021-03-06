# IconUrl element

Specifies the URL of the image that is used to represent your Office Add-in on the Office Ribbon and Office Store. The IconUrl must be an HTTPS address.

## Example

```XML
<OfficeApp>
...
  <Description DefaultValue="Description of your Office Add-in." />
  
  
  <IconUrl DefaultValue="https://contoso.com/assets/icon-32.png" />
  
  
  <HighResolutionIconUrl DefaultValue="https://contoso.com/assets/hi-res-icon.png"/>
...
</OfficeApp>
```


## Attributes

|**Attribute**|**Type**|**Required**|**Description**|
|:-----|:-----|:-----|:-----|
|DefaultValue| URL |Yes|Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](../../reference/manifest/defaultlocale.md) element. |

## Child elements

|  Element | Required | Description  |
|:-----|:-----|:-----|
|  [Override](../../reference/manifest/override.md)   | No | Specifies the setting for additional locale urls |

## Parent element
[OfficeApp]()

## Additional Information

For a mail add-in, the icon is displayed in the  **File** > **Manage add-ins** UI (Outlook) or **Settings** > **Manage add-ins** UI (Outlook Web App). For a content or task pane add-in, the icon is displayed in the **Insert** > **Add-ins** UI. For all add-in types, the icon is also used on the Office Store site, if you publish your add-in to the Office Store.

The image must be in one of the following file formats: GIF, JPG, PNG, EXIF, BMP or TIFF. For content and task pane apps, the image specified must be 32 x 32 pixels. For mail apps, the image must be 64 x 64 pixels. You should also specify an icon for use with Office host applications running on high DPI screens using the [HighResolutionIconUrl](../../reference/manifest/highresolutioniconurl.md) element. For more information, see the section _Create a consistent visual identity for your app_ in [Create effective Office Store apps and add-ins](http://msdn.microsoft.com/library/c66a6e6b-2e96-458f-8f8c-2a499fe942c9%28Office.15%29.aspx).

