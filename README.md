# Macro Tracker

A personal nutrition and activity tracker built on **Google Sheets + Google Apps Script**, with meal estimates from **Claude** (Anthropic API).

Snap a photo of a meal or describe it, and the phone web app asks Claude to estimate calories, protein, carbs, fat and fiber, then logs the result to a Google Sheet. Daily and weekly views compare intake against your targets.

## Features

- **Photo or text logging**: Claude identifies the items, states portion assumptions and returns structured estimates (via tool use)
- **Branded food lookup**: label values per 100 g from [Open Food Facts](https://world.openfoodfacts.org) for named products
- **Corrections**: refine an estimate in plain words ("add 1 tbsp olive oil") before saving
- **Daily dashboard**: totals vs. calorie, protein, carb and fat targets; edit, delete or undo entries; one-tap "log again" for recent meals
- **Activity tracking**: Gym, Dance, Pilates, Self-defense, Long walk, Rest, with notes
- **Averages view**: per-day bars and averages over any range up to 62 days

## Files

| File | What it does |
|---|---|
| `Code.gs` | Server side: Sheets access, Claude API calls, Open Food Facts lookup, saving/editing meals and activity |
| `Index.html` | Mobile web app UI (HTML, CSS, JS in one file) |

## Setup

1. Create a Google Sheet with the tabs **Targets**, **Log**, **Daily** (and optionally **Weekly**).
   - `Targets`: column A labels, column B values. Required labels: `Calorie target (kcal/day)`, `Maintenance calories (kcal/day)`, `Protein min (g/day)`, `Protein max (g/day)`, `Fat (g/day)`, `Carbs (g/day)`.
   - `Log` columns A–M: Date, Time, Meal, Description, Portion, kcal, Protein, Carbs, Fat, Fiber, Confidence, Source, Notes.
   - `Daily`: dates in column A, activity in column K, activity note in column L.
2. In the sheet, open **Extensions → Apps Script** and add `Code.gs` and `Index.html`.
3. Go to **Project Settings → Script properties** and add `ANTHROPIC_API_KEY` with your key (starts with `sk-ant-`). Optional: run `testApiKey` to check it.
4. **Deploy → New deployment → Web app**, execute as yourself, access "Only myself". Open the URL on your phone and add it to the home screen.

To change the model, edit `CONFIG.MODEL` at the top of `Code.gs`.

## Security

The API key lives only in Script Properties and is never stored in the code. Don't commit keys or your web app deployment URL.
