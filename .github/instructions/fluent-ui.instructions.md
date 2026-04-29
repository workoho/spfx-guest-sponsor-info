---
description: >
  Fluent UI v9 coding rules for this repository.
  Apply whenever writing or modifying React components under src/.
applyTo: "src/**"
---

# Fluent UI v9 Rules

This project uses **Fluent UI v9** (`@fluentui/react-components`) exclusively.
The v8 migration is complete — all `@fluentui/react` imports have been removed.

## Import Rules

### Required packages

```typescript
// Components
import { Avatar, Button, Tooltip, ... } from '@fluentui/react-components';
// Icons (tree-shaken SVG)
import { ChatRegular, ChatFilled } from '@fluentui/react-icons';
// Styling
import { makeStyles, tokens, mergeClasses } from '@fluentui/react-components';
// Theme bridge (used in the web part root only)
import { createV9Theme } from '@fluentui/react-migration-v8-v9';
```

### Prohibited

```typescript
// ❌ v8 components — project has fully migrated to v9
import { Persona, Callout, Panel, ActionButton, IconButton,
         Icon, TooltipHost, MessageBar } from '@fluentui/react';
// ❌ v8 icon font
initializeIcons();
// ❌ v8 styling APIs
import { mergeStyles, mergeStyleSets } from '@fluentui/merge-styles';
// ❌ Hardcoded colour values — use tokens instead
const style = { color: '#0078d4' };
// ❌ CSS/SCSS module files — use makeStyles hooks instead
import styles from './Foo.module.scss';
```

## Styling

Use **`makeStyles`** with **`tokens`** for all component-level styles.
Do not add CSS/SCSS module files — all styles live in `makeStyles` hooks.

```typescript
import { makeStyles, tokens, mergeClasses } from '@fluentui/react-components';

const useCardStyles = makeStyles({
  root: {
    backgroundColor: tokens.colorNeutralBackground1,
    color: tokens.colorNeutralForeground1,
    padding: tokens.spacingVerticalM,
  },
  active: {
    borderColor: tokens.colorBrandStroke1,
  },
});

// Usage inside a component:
const styles = useCardStyles();
return <div className={mergeClasses(styles.root, isActive && styles.active)} />;
```

### CSS variable → token mapping

| Old CSS Variable              | Fluent v9 Token                          |
| ----------------------------- | ---------------------------------------- |
| `var(--themePrimary)`         | `tokens.colorBrandForeground1`           |
| `var(--bodyText)`             | `tokens.colorNeutralForeground1`         |
| `var(--neutralLight)`         | `tokens.colorNeutralBackground3`         |
| `var(--neutralQuaternaryAlt)` | `tokens.colorNeutralStroke2`             |
| `var(--neutralSecondary)`     | `tokens.colorNeutralForeground2`         |
| `var(--neutralTertiary)`      | `tokens.colorNeutralForeground3`         |
| `var(--neutralLighter)`       | `tokens.colorNeutralBackground2`         |
| `var(--white)`                | `tokens.colorNeutralBackground1`         |
| `var(--link)`                 | `tokens.colorBrandForegroundLink`        |
| `var(--successText)`          | `tokens.colorPaletteGreenForeground1`    |

## Theme integration

`FluentProvider` is already set up in the web part root — do **not** add a second one inside
components. Theming flows down through React context automatically.

```tsx
// Already in GuestSponsorInfoWebPart.ts — do NOT duplicate:
const v9Theme = theme ? createV9Theme(theme) : undefined;
return <FluentProvider theme={v9Theme}>{children}</FluentProvider>;
```

## Component equivalents (v8 → v9)

| v8                               | v9                                                         |
| -------------------------------- | ---------------------------------------------------------- |
| `Persona` (avatar display)       | `Avatar` with `color="colorful"`                           |
| `PersonaPresence` enum           | `PresenceBadge` `status` string prop                       |
| `Callout`                        | `Popover` + `PopoverTrigger` + `PopoverSurface`            |
| `Panel`                          | `OverlayDrawer` + `DrawerHeader` + `DrawerBody`            |
| `ActionButton` / `IconButton`    | `Button` with `appearance="subtle"`                        |
| `Icon iconName="Chat"`           | `<ChatRegular />` (SVG from `@fluentui/react-icons`)       |
| `TooltipHost`                    | `Tooltip` with `relationship="label"`                      |
| `Link`                           | `Link` from `@fluentui/react-components`                   |
| `MessageBar` + `MessageBarType`  | `MessageBar` + `MessageBarBody` with `intent` prop         |
| `IButtonStyles` (inline styles)  | `makeStyles` with `tokens`                                 |

## Icon pattern

Use `bundleIcon` to combine filled + regular variants (enables filled-on-hover in Fluent):

```typescript
import { bundleIcon, ChatRegular, ChatFilled } from '@fluentui/react-icons';

const ChatIcon = bundleIcon(ChatFilled, ChatRegular);
// Then: <Button icon={<ChatIcon />} />
```

### MDL2 → @fluentui/react-icons name mapping

| MDL2 name    | v9 import               |
| ------------ | ----------------------- |
| `Chat`       | `ChatRegular`           |
| `Mail`       | `MailRegular`           |
| `Phone`      | `CallRegular`           |
| `CellPhone`  | `PhoneRegular`          |
| `CityNext`   | `BuildingRegular`       |
| `MapPin`     | `LocationRegular`       |
| `Copy`       | `CopyRegular`           |
| `Accept`     | `CheckmarkRegular`      |
| `Org`        | `OrganizationRegular`   |
