# PatternHub

**PatternHub** is a structured collection of Adobe Illustrator patternmaking scripts — organized by **author → block → version** — with both stable and beta releases.  
Each kit contains ExtendScript (`.jsx`) files that automate professional pattern drafting directly inside Illustrator.

---

## Project Structure

PatternHub/  
├─ README.md  
├─ manifest.json  
├─ .gitignore  
├─ .gitattributes  
├─ .editorconfig  
├─ /MiniTools/ — mini-scripts that makes general drafting quicker
├─ /Aldrich/ — Aldrich patternmaking scripts  
│ ├─ Bodice/  
│ ├─ Trousers/  
│ ├─ Skirt/  
│ └─ Sleeve/  
├─ /Muller/  
└─ /Hofenbitzer/

---

## Quick Start

1. **Open Adobe Illustrator**
2. Go to **File → Scripts → Other Script…**
3. Browse to a `.jsx` file inside any version folder (e.g. `Aldrich/Bodice/v1.0.0/bodice.jsx`)
4. Run the script to generate the pattern.

Each block version includes a `metadata.json` file describing measurements, layers, color settings, and label rules.

---

## Versioning

PatternHub follows **Semantic Versioning**

- `1` — Stable release
- `1.beta` — Beta release (testing stage)

**Branches**

- `main` → stable releases only
- `develop` → work in progress (betas)

---

## MiniScripts

Lightweight Illustrator utility scripts (e.g., add lines, add notches, measure paths).  
They live in `/MiniScripts/` and are versioned separately inside the manifest.

---

## Manifest

The root `manifest.json` acts as the registry of all available authors, blocks, and versions.  
Each pattern folder contains its own `metadata.json` file for local details.

---

## Contributing

1. Fork or clone the repo.
2. Add a new author or block folder in the correct structure.
3. Update both `manifest.json` and any local `metadata.json`.
4. Commit and push to the `develop` branch.

**Formatting standards**

- Indentation: 2 spaces
- End-of-line: LF
- Do not commit zip or build files.

---

## License

This project is licensed under the **Creative Commons Attribution-NonCommercial 4.0 International License (CC BY-NC 4.0)**.  
You may share and adapt the material for non-commercial purposes, provided that appropriate credit is given.  
Read the full license here: [CC BY-NC 4.0](https://creativecommons.org/licenses/by-nc/4.0/)

---

## Credits

Created by the **more than patterns** team —  
inspired by patternmaking systems such as _Aldrich_, _Müller & Sohn_, and _Hofenbitzer_.

## Communities

- [Discord](https://discord.gg/CFvfXfZUTa)
- [Reddit](https://www.reddit.com/r/morethanpatterns/)
- [Website](https://morethanpatterns.com/)
