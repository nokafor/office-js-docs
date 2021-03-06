
# SourceLocation element

Specifies the source file location(s) for your Office Add-in as a URL between 1 and 2018 characters long. The source location must be an HTTPS address, not a file path.

## Example

```XML

<OfficeApp>
...
  <DefaultSettings>
    <SourceLocation DefaultValue="https://contoso.com/" />
  </DefaultSettings>
...
</OfficeApp>

```

## Attributes

|**Attribute**|**Type**|**Required**|**Description**|
|:-----|:-----|:-----|:-----|
|DefaultValue|URL|Yes|Specifies the default value for this setting for the locale specified in the [DefaultLocale](../../reference/manifest/defaultlocale.md) element.|


## Child elements


|  Element | Required | Description  |
|:-----|:-----|:-----|
|  [Override](../../reference/manifest/override.md)   | No | Specifies the setting for additional locale urls |

## Parent element

[DefaultSettings](../../reference/manifest/defaultsettings.md) (Content and task pane add-ins)

[FormSettings](../../reference/manifest/formsettings.md) (Mail add-ins)


