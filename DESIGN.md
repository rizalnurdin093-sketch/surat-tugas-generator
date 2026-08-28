# DESIGN.md — Surat Tugas Generator Design System

> Single source of truth untuk visual & interaction design. Referensi ini digunakan untuk konsistensi & revert capability.

---

## 1. Design Philosophy

**Corporate Clean** — Putih dominan, 1 warna aksen (Teal), tipografi rapi, proporsional, minimal visual noise. Mirip Linear/Notion untuk internal tool. Tidak "AI slop": tidak gradient berlebihan, tidak shadow tebal, tidak radius berlebihan, tidak warna random.

**Prinsip:**
- **Hierarchy via weight & space**, bukan warna berlebihan
- **8px base grid** untuk semua spacing
- **System font stack** — native, cepat, familiar
- **Restrained radius** — 8px card, 6px input, 4px button
- **Subtle shadow** — elevation hanya untuk card & dropdown

---

## 2. Color Palette

| Token | Hex | Usage |
|-------|-----|-------|
| `--color-primary` | `#00897B` | Primary actions, links, focus ring, active icon |
| `--color-primary-hover` | `#00796B` | Hover state primary |
| `--color-primary-light` | `#E0F2F1` | Subtle bg for active nav, badges |
| `--color-bg` | `#FAFAFA` | Page background |
| `--color-surface` | `#FFFFFF` | Card, modal, dropdown bg |
| `--color-border` | `#E2E8F0` | Input border, table border, divider |
| `--color-border-hover` | `#CBD5E1` | Hover border |
| `--color-text` | `#1E293B` | Primary text (slate-800) |
| `--color-text-secondary` | `#64748B` | Secondary text (slate-500) |
| `--color-text-muted` | `#94A3B8` | Placeholder, disabled, caption (slate-400) |
| `--color-error` | `#DC2626` | Error text, danger actions |
| `--color-error-bg` | `#FEF2F2` | Error alert bg |
| `--color-success` | `#059669` | Success text |
| `--color-success-bg` | `#ECFDF5` | Success alert bg |
| `--color-focus` | `#0D9488` | Focus ring (teal-600) |

---

## 3. Typography

**Font Stack:** `system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif`

| Style | Size | Line-height | Weight | Usage |
|-------|------|-------------|--------|-------|
| `--text-display` | 28px / 1.75rem | 1.2 | 600 | Page title (h1) |
| `--text-heading` | 22px / 1.375rem | 1.3 | 600 | Section title (h2) |
| `--text-body` | 15px / 0.9375rem | 1.5 | 400 | Body, form labels, table cells |
| `--text-body-sm` | 13px / 0.8125rem | 1.4 | 400 | Helper text, table header |
| `--text-caption` | 12px / 0.75rem | 1.4 | 400 | Timestamp, badge, footer |

---

## 4. Spacing Scale (8px base)

| Token | Value | Usage |
|-------|-------|-------|
| `--space-xs` | 4px | Icon gap, tight spacing |
| `--space-sm` | 8px | Form group gap, inline gap |
| `--space-md` | 16px | Card padding, section gap |
| `--space-lg` | 24px | Page section gap, card gap |
| `--space-xl` | 32px | Major section gap |
| `--space-2xl` | 48px | Page top/bottom padding |

---

## 5. Border Radius

| Token | Value | Usage |
|-------|-------|-------|
| `--radius-sm` | 4px | Button, badge, avatar |
| `--radius-md` | 6px | Input, select, dropdown |
| `--radius-lg` | 8px | Card, table container |
| `--radius-xl` | 12px | Modal, popover |

---

## 6. Shadows

| Token | Value | Usage |
|-------|-------|-------|
| `--shadow-sm` | `0 1px 2px rgba(15,23,42,0.05)` | Hover row, subtle |
| `--shadow-md` | `0 4px 6px -1px rgba(15,23,42,0.1), 0 2px 4px -1px rgba(15,23,42,0.06)` | Card, dropdown |
| `--shadow-lg` | `0 10px 15px -3px rgba(15,23,42,0.1), 0 4px 6px -2px rgba(15,23,42,0.05)` | Modal |

---

## 7. Icon System

- **Inline SVG**, 20×20 (nav), 18×18 (inline), 24×24 (primary action)
- **Stroke width:** 2px (outline style)
- **CurrentColor** — inherit text color
- Sumber: Heroicons outline (manual copy, no dep)

---

## 8. Components

### Navbar
- Height: 64px
- Sticky top, z-index 50
- BG: `--color-surface` + bottom border `--color-border`
- Brand: text "SUGEN", weight 600, size 1.125rem, color `--color-text`
- Nav icons: 20×20, `--color-text-secondary` default, `--color-primary` active/hover
- Profile avatar: 32×32, initials, bg `--color-primary-light`, color `--color-primary`
- Dropdown: `--shadow-lg`, `--radius-lg`, min-width 200px

### Card
- BG: `--color-surface`
- Border: 1px solid `--color-border`
- Radius: `--radius-lg` (8px)
- Padding: `--space-lg` (24px)
- Shadow: `--shadow-sm` (barely visible)

### Button
- **Primary:** BG `--color-primary`, text white, `--radius-md`, padding `10px 20px`, font `--text-body`, weight 500
- **Primary hover:** `--color-primary-hover`
- **Secondary (ghost):** Transparent BG, text `--color-text-secondary`, border `--color-border`
- **Secondary hover:** BG `--color-bg`, text `--color-text`
- **Danger:** BG `--color-error`, text white
- Focus: `0 0 0 3px var(--color-primary-light)`

### Input
- Width: 100%, padding `10px 12px`, `--radius-md`, border `--color-border`
- Font: `--text-body`
- Placeholder: `--color-text-muted`
- Hover border: `--color-border-hover`
- Focus: border `--color-primary`, ring `0 0 0 3px var(--color-primary-light)`

### Table
- Border-collapse: collapse
- Header: BG `--color-bg`, text `--text-body-sm` uppercase tracking-wide, color `--color-text-secondary`
- Cell: padding `12px 16px`, border-bottom `--color-border`
- Row hover: BG `--color-bg`
- Font: `--text-body`

### Flash/Alert
- Padding: `12px 16px`, `--radius-md`
- Error: BG `--color-error-bg`, text `--color-error`, border `--color-error` 20%
- Success: BG `--color-success-bg`, text `--color-success`, border `--color-success` 20%

---

## 9. Layout

- **Max content width:** 960px (slightly wider than before for table breathing room)
- **Page padding:** `--space-lg` (24px) mobile, `--space-xl` (32px) desktop
- **Navbar + Main + Footer** — footer sticky bottom if content short

---

## 10. Interaction States

| State | Visual |
|-------|--------|
| Hover (nav icon) | Color `--color-primary`, BG `--color-primary-light` |
| Active (nav) | Color `--color-primary`, indicator dot |
| Focus visible | Ring `--color-focus` 3px |
| Disabled | Opacity 0.5, cursor not-allowed |
| Loading button | Spinner, disable pointer |

---

## 11. Responsive Breakpoints

| Breakpoint | Width | Adjustments |
|------------|-------|-------------|
| Mobile | < 640px | Stack form, smaller padding, table horizontal scroll |
| Tablet | 640–1024px | Normal |
| Desktop | > 1024px | Max-width 960px centered |

---

## 12. Migration Notes (from old design)

**Removed:**
- Green primary (`#27ae60`) → replaced by Teal (`#00897B`)
- Dark navbar (`#2c3e50`) → Light surface navbar
- Text links in nav → Icon-only nav
- Large radius (8px→8px same but more consistent)
- Inline styles in templates → All in CSS

**Added:**
- Profile dropdown
- Icon-only navigation
- Consistent design tokens
- Better visual hierarchy

---

## 13. Usage Rules

1. **Never add new colors** — use tokens only
2. **Never hardcode spacing** — use scale
3. **Never add new fonts** — system stack only
4. **Icon size** — 20×20 default, 24×24 for primary CTA
5. **One primary action per view** — others secondary/ghost

---
*File ini single source of truth. Edit di sini, lalu apply ke CSS/templates.*