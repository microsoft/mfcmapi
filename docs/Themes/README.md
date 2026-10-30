# Optional theme files

This directory contains optional theme files. Nothing here is built in or loaded
by default. To try them, open **Tools > Themes**, choose this directory, enable
color themes, and select a theme. Other theme files are provided alongside the
examples below.

## Writing a theme

Start with [Halloween](Halloween.json) or [Christmas](Christmas.json):

- Halloween is a complete example with explicit colors for both appearances.
- Christmas is a sparse example: it overrides selected colors and inherits the
  rest from Standard.

1. Copy an example to a new `.json` file, saved as UTF-8.
2. Give it a unique `id` and your own display `name`.
3. Edit the colors in both the `light` and `dark` palettes.
4. Change the `schedule`, or remove it for a manual-only theme.
5. In **Tools > Themes**, choose the directory containing your file and enable
   color themes. Select your theme to try it.

After editing a file, click **Reload themes** to see your changes. Keep your files
directly in the chosen directory, not in subdirectories.

## Theme information

Keep `schemaVersion` set to `1` and use the field names exactly as shown in the
examples. JSON field names and theme IDs are case-sensitive.

| Field | Meaning |
|---|---|
| `id` | Required unique identifier, such as `my.spring`. Keep it unchanged when updating your theme. Do not reuse a built-in theme's ID. |
| `name` | Required display name shown in the theme list. |
| `colors` | Required `light` and `dark` objects. Include only the color entries you want to customize. |
| `schedule` | Optional dates when your theme should be selected automatically. |

## Colors

Each palette role is an object:

```json
{ "value": "DarkGreen", "system": -1, "highContrast": 8 }
```

- `value` is a CSS color name, such as `DarkGreen`, or a six-digit hex color,
  such as `#228B22`. Names and hex digits are case-insensitive; both `gray` and
  `grey` spellings work. Transparency and CSS expressions are not supported.
- `system` optionally uses a Windows system color instead of `value`. Omit it
  or set it to `-1` to use your chosen color.
- `highContrast` chooses the Windows system color to use in High Contrast mode.
  Keep the example's mappings unless you have a reason to change them.

Omitted color entries inherit from Standard's matching light or dark palette,
including its system-color and High Contrast mappings. Use an empty object (`{}`)
to inherit an entire palette. You can omit the bitmap and `Test*` entries.
An entry you supply still needs `value` and, for visible colors, `highContrast`;
individual fields within that entry are not inherited.

### Disabled menu items

Menus use different palette roles for enabled and disabled commands. `Text` is
the enabled command color; `TextDisabled` is used when a command is disabled.
When the pointer highlights a command, enabled text uses `GlowText` on
`SelectedBackground`, while a disabled command uses `TextDisabledSelected` on
`DisabledSelectedBackground` with a one-pixel `DisabledSelectedBorder` outline.

Disabled commands should be recognizable as disabled, but still easy to read.
Choose disabled colors that are visibly different from their enabled
counterparts; do not just repeat `Text` or `GlowText`. For example:

```json
{
  "Text": { "value": "#202020", "highContrast": 8 },
  "TextDisabled": { "value": "#606060", "highContrast": 17 },
  "GlowText": { "value": "#FFFFFF", "highContrast": 14 },
  "TextDisabledSelected": { "value": "#D2D2D2", "highContrast": 17 }
}
```

These are role examples, not a complete palette. Choose values based on your
own `Background` and `SelectedBackground` colors and check both appearances.

### Recommended contrast

These are recommendations for custom themes, not restrictions imposed when a
theme file is loaded. Contrast is the ratio between two colors, from `1:1`
(identical) to `21:1` (black and white). Check every pairing in both palettes.

- Aim for at least **4.5:1** between text and its background, including normal,
  read-only, selected, hover, status, pane-header, and text-highlight text.
- For `TextDisabled`, aim for at least **3:1** against `Background`. For
  `TextDisabledSelected`, aim for at least **4.5:1** against `SelectedBackground`.
- Make enabled and disabled text clearly different: aim for at least **3:1**
  between `Text` and `TextDisabled`, and **3:1** between `GlowText` and
  `TextDisabledSelected`.
- Aim for at least **3:1** between the background and visual indicators such as
  selection, hover, text-highlight fills, and the focus indicator.

The bundled themes are checked against these targets by unit tests. Custom
themes can use different values, but low contrast or nearly identical enabled
and disabled colors can make text difficult to read or state changes hard to
notice.

If you change system-color mappings, use these numbers, not RGB values:

| Number | Windows system color |
|---|---|
| `5` | Window background |
| `8` | Window text |
| `13` | Selection background |
| `14` | Selection text |
| `15` | Button background |
| `17` | Disabled text |
| `18` | Button text |

Other Windows system-color numbers from 0 to 31 are accepted except 25.
Visible color entries require a `highContrast` mapping; the bitmap and `Test*`
entries do not.

## Schedules

```json
"schedule": {
  "kind": "dateRange",
  "start": "10-25",
  "end": "10-31",
  "priority": 10,
  "seasonal": false
}
```

Enable **Use scheduled theme** in the Themes dialog to use these dates.
Selecting a theme yourself turns scheduling off.

- `start` and `end` use `MM-DD` and include both dates, repeating every year.
  A range can cross New Year's Day. February 29 uses February 28 in non-leap years.
- `priority` defaults to `0`. Higher numbers win when schedules overlap.
  For equal priorities, the shorter date range wins, then the theme ID breaks ties.
- `seasonal` defaults to `false`. Set it to `true` for seasons that should shift
  six months in the Southern Hemisphere, based on the Windows country/region
  setting. Leave it `false` for holidays and other fixed-date events.

For example, to make your own Spring theme take precedence over built-in Spring,
use a unique ID such as `my.spring`, dates `03-01` through `05-31`, priority `10`,
and `seasonal: true`. Built-in Spring remains available in the theme list.

## Fixing errors

If your file appears as **in error**, select it and read **Theme file errors**.
Check for invalid JSON, misspelled field or color names, incomplete color entries,
or a duplicate ID. Fix the file and click **Reload themes**.

A broken file does not prevent other themes from loading.
