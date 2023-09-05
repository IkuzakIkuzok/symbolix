
# Symbolix

Hyphen, or not hyphen, that is the question

## Overview

Symbolix is a Add-in for Microsoft Word to convert hyphens into minus signs, en-dashes, or em-dashes.

## Requirements

- Microsoft Word 2016 or later
- .NET Framework 4.7.2 runtime
- Visual Studio 2010 Tools for Office runtime

See [MSDN](https://learn.microsoft.com/en-us/visualstudio/vsto/how-to-install-the-visual-studio-tools-for-office-runtime-redistributable?view=vs-2022) for more details.

## Usage

1. Install Symbolix to Microsoft Word by double-clicking `symbolix.vsto` file.
1. Write or open a document in Microsoft Word.
1. In the `Run` section in `Symbolix` tab, click `This document`, `Selection`, or `All documents` button.

## Features

- Convert hyphens before numeric characters into minus signs (e.g., from "-123" (hyphen) to "−123" (minus)).
- Simple word replacement.
- Convert hyphens into minus signs, en-dashes, or em-dashes (e.g., from "current density-voltage" (hyphen) to "current density–voltage" (en-dash)).

All changes made by Symbolix are tracked by Microsoft Word, so you can undo them by rejecting them.

## Settings

You can specify the behabiou of Symbolix globally or for each folders.
All settings are written in `.symbolixconfig` file in JSON format.
The schema file for `.symbolixconfig` is available as `config-schema.json`.

### Properties

#### `save`: string[]

Specify when to save the target file.
You can specify `"before"` and/or `"after"`, which mean save the target file before and/or after the execution, respectively.

#### `check-minus`: boolean

Indicate whether to check the minus sign before numeric character.

#### `replace`: object[]

Specifies words for simple substitutions.
Each object has the following properties:

| Name | Required | Type | Description |
| ---- | -------- | ---- | ----------- |
| `action` | Yes | one of `"add"` or `"remove"` | `"add"` for adding the word to the substitution list, `"remove"` for removing the word from the substitution list. |
| `find` | Yes | string | The word to be substituted. |
| `replace` | Yes if `action` is `"add"`; otherwise, no | string | The word to substitute. |

#### `minus`, `en-dash`, and `em-dash`: object[]

Specifies words to substitute hyphen into minus, en-dash, or em-dash.
Each object has the following properties:

| Name | Required | Type | Description |
| ---- | -------- | ---- | ----------- |
| `action` | Yes | one of `"add"` or `"remove"` | `"add"` for adding the word to the substitution list, `"remove"` for removing the word from the substitution list. |
| `find` | Yes | string | The word to be substituted. |

All hyphen signs in the `find` property are substituted into the corresponding sign.

For example, if you added "current density-voltage" to the en-dash list, this word in the document is converted into "current density–voltage" (en-dash).

### Default settings

Default settings are equivalent to the following `.symbolixconfig` file:

```json
{
  "save": [ "after" ],
  "check-minus": true,
  "replace": [],
  "minus": [],
  "en-dash": [],
  "em-dash": []
}
```

### Settings resolving rule

Symbolix resolves settings in the following order:

1. Global settings: `C:\Users\{UserName}\Documents\.symbolixconfig`
1. Parent folder settings: `.symbolixconfig` in the parent folder of the target file. Note that Symbolix searches `.symbolixconfig` in the parent folder recursively.
1. Folder specific settings: `.symbolixconfig` file in the same folder as the target file.

If there are multiple settings, Symbolix merges them in the order above, and the latter settings override the former settings.
For the parent folder settings, the settings in the folder closer to the target file override the settings in the folder farther from the target file.

If the target file is not in any folder, Symbolix uses only global settings.

## License

Symbolix is licensed under the MIT license. See the [LICENSE file](https://github.com/IkuzakIkuzok/symbolix/blob/main/LICENSE) for details.
