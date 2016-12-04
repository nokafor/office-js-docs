# Additions to Office.context

Below section describes the planned additions to Office.Context object in order to provide the add-in environment information to developers. Please review and provide your feedback. One of the best ways of providing your input is by opening new issue in GitHub using the links available below.

_**Note**: below listed features are still under design and review phase and hence not yet available as part of the product. The final design is subject to change. Once the feature is made available, the final specification will be published as part of the master repository._

## Host and platform information 

`Office.context` object represents the runtime environment of the add-in. Among other things it contains Office theme, touch enabled flag, display language, etc. We are making two additions to this object by introducing `host` and `platform` information.

### Platform 
This could be accessed using:  

`var host = Office.context.platform;`

It returns a string whose possible values could be one of the following: 
* "PC" (Windows desktop environment) 
* "OFFICE_ONLINE" (Office online environment) 
* "MAC" (Office on Mac)
* "IOS" (Office on iPad)
* `null`: If the site is not running within a Office host (such as Excel or Word), then `null` value is returned. 

The following enumratons could also be used to check the value being retured: 

```js
switch (Office.context.platform) {
                case Office.PlatformTypes.PC:
                   // do something
                case Office.PlatformTypes.OFFICE_ONLINE:
                   // do something
}

```

### Host

This could be accessed using:  

`var platform = Office.context.host;`

It returns a string whose possible values could be one of the following: 
* "EXCEL" 
* "ONENOTE"
* "OUTLOOK"
* "POWERPOINT"
* "WORD"
* "PROJECT"
* "ACCESS"
* "HOST"

The following enumratons could also be used to check the value being retured: 

```js
	switch (Office.context.host) {
                case Office.HostTypes.EXCEL:
                   // do something
                case Office.HostTypes.ONENOTE:
                   // do something
	}
```



## Office diagnostic information 
Provides diagnostic information about Office add-in, which could be used to collect diagnostic information about the Office tuntime environment. This could be accessed using: 

`var diagnostics = Office.context.diagnostics;`

It returns an JSON object whose structure is as follows: 

```json
{
	"host": "..",
	"platform": "..",
	"version": ".."
}
```

Please refer above for host and platform values. For version, we'll make the best effort to return the version number of the Office specific to the platform under which host is running. 

## Give feedback

We need it, you want to give it. Feedback is much easier to give now that we're on GitHub. Check out the docs and let us know your feedback by submitting [issues](https://github.com/OfficeDev/office-js-docs/issues) directly in this repository.

For API support, you can post questions to the community on [StackOverflow](http://stackoverflow.com/questions/tagged/office-js) and tag them with [office-js].
