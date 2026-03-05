# Competition & Links Admin Guide

## Add or update competition tabs (Metrics Admin)
1. Open **Metrics Admin**.
2. Add/edit rows in the category table.
3. Save categories.
4. Enabled categories become tabs immediately (no redeploy).

### Competition_Categories schema
- `CategoryKey` (stable internal key)
- `CategoryName` (tab label)
- `Enabled` (TRUE/FALSE)
- `SortOrder` (number)
- `Goal` (optional)
- `Notes` (optional)
- `UpdatedAt`

### Competition_Entries schema
- `Date`
- `Store`
- `CSR`
- `CategoryKey`
- `Value`
- `Notes`
- `UpdatedAt`

## Patio Cushion Signup link updates (Links Admin)
1. Open **Links Admin**.
2. Edit `Label` and `URL` for `PATIO_CUSHION_SIGNUP`.
3. Save links.
4. Dashboard + Patio Signups tab update immediately (no redeploy).

### App_Links schema
- `Key`
- `Label`
- `Url`
- `UpdatedAt`
