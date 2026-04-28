# PMI Isometric Plotter

This workspace now includes two ways to work:

- `index.html` is a small web app for drawing directly on branded isometric paper.
- `isometric_plotter.py` is the original PDF generator if you still want printable output from JSON.

## Web app

Open `index.html` in a browser, or serve the folder locally:

```bash
/Users/thisaruperera/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 -m http.server 8000
```

Then open `http://localhost:8000`.

### Web app features

- snap-to-grid drawing on an isometric paper layout
- PMI branding in the lower-left corner instead of the source footer
- undo, clear, and PNG export
- live guide line while placing each segment

## PDF script

Run the PDF version with:

```bash
/Users/thisaruperera/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 isometric_plotter.py --spec example_plot.json --output branded_isometric_plot.pdf
```
