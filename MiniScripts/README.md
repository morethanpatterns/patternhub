# MiniScripts

**MiniScripts** are small, reusable Adobe Illustrator ExtendScript utilities that assist with pattern drafting, measurement, and alignment.

---

## Structure

MiniScript/  
├─ README.md  
├─ CHANGELOG.md  
└─ src/  
 ├─ add_and_label_line.jsx  
 ├─ add_guides.jsx  
 ├─ measure_along_path.jsx  
 └─ utils/geometry.jsx

---

## Available Scripts

| File                     | Function                                      |
| ------------------------ | --------------------------------------------- |
| `add_line.jsx`           | Adds a straight line at specified coordinates |
| `add_guides.jsx`         | Creates guide lines at defined positions      |
| `measure_along_path.jsx` | Measures a given distance along a path        |
| `utils/geometry.jsx`     | Shared geometric helper functions             |

---

## Usage

In another script:

```javascript
#include "../MiniScripts/src/add_line.jsx"
```
