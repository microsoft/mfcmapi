# Optional theme files

This directory contains optional theme files. Nothing here is built in or loaded
by default. To try them, open **Tools > Themes**, choose this directory, enable
color themes, and select a theme. Other theme files are provided alongside the
examples below.

**Ctrl+Shift+T** selects the next theme; **Ctrl+Shift+R** selects the previous
theme. Both wrap through the theme list and turn scheduling off. Color themes
must be enabled. **Ctrl+Shift+D** toggles light/dark appearance.

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

Colors never inherit from another role, even if two Standard entries happen to
have identical values. Standard's appearance remains unchanged.

### Title bars, column headers, and menus

These surfaces have independent roles in each palette:

| Surface | Roles |
|---|---|
| Title bar | `TitleBarBackground`, `TitleBarText`, `TitleBarButton`, `TitleBarButtonHover` |
| Column headers | `ColumnHeaderBackground`, `ColumnHeaderText`, `ColumnHeaderArrow` |
| Menu bar and popup menus | `MenuBackground`, `MenuText`, `MenuTextDisabled` |

`TitleBarButton` colors the close, maximize/restore, and minimize glyphs.
Highlighted caption buttons use `TitleBarButtonHover` on `GlowBackground`.
Column-header colors cover both the labeled columns and unused header area.
`MenuGutter` still controls the popup icon/checkmark gutter; menu selections keep
their existing selection roles.

Omitting any of these roles inherits only that role's Standard entry, not the
theme's `Background`, `Text`, `Arrow`, or any other customized color. Existing
theme files remain valid. The preview includes these surfaces. Built-in seasons
use selective accents. Every bundled theme file explicitly defines all ten
surface roles in both appearances, with coordinated title-bar accents,
column-header colors, and menu backgrounds. Title-bar accents can be selected
separately from status-bar colors. Pride Rainbow keeps the window restrained:
off-white or charcoal work areas, purple title bars, orange menu gutters, faint
violet header tints, blue selection highlights, and muted green and gold interaction accents. Its saturated
rainbow is concentrated in the status gradient. Christmas pairs a green title
bar with its red status bar. These are independent literal color values, not
references to other roles.

Selected list rows paint `GlowText` on `SelectedBackground`; hovered rows paint
`GlowText` on `GlowBackground`. Both paint the matching pair together rather
than mixing theme text with a native Windows highlight background.
Tree notification markers use the row's text color when selected or hovered,
keeping their normal accent color when unhighlighted.
Submenu arrows use the menu item's text color, including selected and disabled states.

### Optional multicolor status bar

Pride Rainbow tries a full-height horizontal rainbow gradient. Status messages
use `StatusText` with a thin, one-pixel `StatusTextOutline` directly on the gradient,
without solid backplates. The outline defaults to Standard's explicit black entry
in both appearances; it never inherits from another role. The preview uses the
same treatment. Other bundled themes, including Standard, remain unchanged.

To opt in, add a top-level `statusBar` object (not inside `colors`):

```json
"statusBar": {
  "style": "gradient",
  "colors": ["#E40303", "#FF8C00", "#FFED00", "#008026", "#004DFF", "#750787"]
}
```

Change `style` and reload themes to compare:

- `gradient`: smooth horizontal transitions between evenly spaced color stops,
  filling the whole status bar and its bottom/side borders.
- `bands`: equal-width, crisp color blocks filling that same area.
- `ribbon`: crisp bands along the top edge, up to five pixels high, leaving room
  for text on the original status background.

Colors are explicit CSS names or six-digit hex values, ordered left to right.
Up to 16 colors are accepted. A single color produces a solid fill; an empty
array disables the decoration. Colors are shared between appearances, and High
Contrast disables the decoration and outlined text treatment.

Omitting the decoration inherits Standard's undecorated status bar, never
colors from another role. The original top-level `statusRibbon` array is still
accepted as shorthand for `ribbon`; do not specify both forms.
For now, only Pride Rainbow opts in so its appearance can be evaluated before
extending the treatment to other themes.

### Disabled menu items

Menus use different palette roles for enabled and disabled commands. `MenuText`
is the enabled command color; `MenuTextDisabled` is used when a command is disabled.
When the pointer highlights a command, enabled text uses `GlowText` on
`SelectedBackground`, while a disabled command uses `TextDisabledSelected` on
`DisabledSelectedBackground` with a one-pixel `DisabledSelectedBorder` outline.

Disabled commands should be recognizable as disabled, but still easy to read.
Choose disabled colors that are visibly different from their enabled
counterparts; do not just repeat `MenuText` or `GlowText`. For example:

```json
{
  "MenuText": { "value": "#202020", "highContrast": 8 },
  "MenuTextDisabled": { "value": "#606060", "highContrast": 17 },
  "GlowText": { "value": "#FFFFFF", "highContrast": 14 },
  "TextDisabledSelected": { "value": "#D2D2D2", "highContrast": 17 }
}
```

These are role examples, not a complete palette. Choose values based on your
own `MenuBackground`, `SelectedBackground`, and `DisabledSelectedBackground`
colors and check both appearances.

### Recommended contrast

These are recommendations for custom themes, not restrictions imposed when a
theme file is loaded. Contrast is the ratio between two colors, from `1:1`
(identical) to `21:1` (black and white). Check every pairing in both palettes.

- Aim for at least **4.5:1** between text and its background, including normal,
  read-only, selected, hover, status, pane-header, title-bar, column-header,
  menu, and text-highlight text.
- For `MenuTextDisabled`, aim for at least **3:1** against `MenuBackground`. For
  `TextDisabledSelected`, aim for at least **4.5:1** against `DisabledSelectedBackground`.
- Make enabled and disabled text clearly different: aim for at least **3:1**
  between `MenuText` and `MenuTextDisabled`, and **3:1** between `GlowText` and
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
