# Cycle Calendar Planner

A single-page week planner for i.c.stars Milwaukee Tech Foundations cycles. Drag workshops
onto the week grid, group them by category, and adjust the schedule in the browser.

**Live site:** https://icstars-milwaukee.github.io/cycle-calendar-planner/

## How it works

The whole app is one self-contained `index.html` — no build step, no dependencies, no server.
It runs entirely in the browser.

## Where your data lives

**All data is stored locally in your own browser** via `localStorage`. Nothing is uploaded,
and no account or network connection is required after the page loads.

Practical consequences:

- Your plan is tied to the browser and device you created it on — it will not follow you to
  another machine or another browser.
- Clearing site data or browsing history for this domain erases the saved plan.
- Private/incognito windows discard the plan when the window closes.

## Running it locally

Clone the repo and open `index.html` in a browser. That's it — no server needed.

```
git clone https://github.com/icstars-milwaukee/cycle-calendar-planner.git
```

## Deployment

GitHub Pages serves this repo from the `main` branch root. A `.nojekyll` file is present so
Pages publishes the files as-is rather than running them through Jekyll.
