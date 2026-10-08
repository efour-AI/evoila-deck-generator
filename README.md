# evoila deck generator

Builds an evoila walking deck (PPTX + PDF) for one row in the Airtable **Deck Requests** table
(evoila Content Portal base) and writes both files back to the row.

Flow:

1. Someone submits the Deck Requests form (audience, framing, prospect, date, presenters, meeting notes).
2. Airtable automation "Deck request: plan and build" asks Airtable AI to pick slides and write talking points,
   saves the plan to the row, sets Status = Building and calls this repo's `repository_dispatch` (event `build-deck`).
3. `.github/workflows/build-deck.yml` runs `run_request.py`: downloads the master template and the engine zip
   from the **Deck Assets** table, builds the deck, puts the talking points in the speaker notes,
   converts to PDF with LibreOffice and uploads both files to the row. Status = Ready (or Error with detail).
4. Airtable automation "Deck ready: email requester" emails the PPTX and PDF and sets Status = Sent.

Repo secret needed: `AIRTABLE_TOKEN` (Airtable personal access token with `data.records:read` and
`data.records:write` on the evoila Content Portal base).

Slide content lives in `build_library.py` inside the engine zip on the Deck Assets table, not in this repo.
To rebuild a deck by hand: Actions tab, "Build walking deck", Run workflow, paste the record ID.
