# NFPA 92 / 92B Smoke Management Systems – Calculator & Compliance Checklist

**SI units** · Algebraic methods for malls, atria and large-volume spaces

Streamlit application providing:

1. **Preliminary design calculations** for smoke management systems (axisymmetric plume, exhaust rate, layer depth, filling time, inlet count)
2. A comprehensive **data checklist** aligned with NFPA 92 design-documentation expectations
3. **Excel report** export

## Standard status

- **NFPA 92B** (*Guide for Smoke Management Systems in Malls, Atria, and Large Areas*) was withdrawn in the 2011 cycle and incorporated into **NFPA 92** (*Standard for Smoke Control Systems*).
- The algebraic equations and design approaches implemented here follow the classic NFPA 92B methods retained in NFPA 92 Chapter 5, expressed entirely in **SI units**.

## Features

### Calculations tab
- Geometry of the large-volume space (length, width, height, highest walking surface)
- Design fire: steady-state or t² growth (slow / medium / fast / ultra-fast)
- Convective fraction and ambient conditions
- Design smoke-layer interface height with automatic checks:
  - Clear height ≥ 1.83 m above highest walking surface
  - Layer depth ≥ 20 % of floor-to-ceiling height (or note for engineering justification)
- Axisymmetric plume mass-flow rate and required volumetric exhaust
- Approximate smoke temperature and density
- Makeup-air quantity (85–95 % of exhaust)
- Plugholing screening → suggested minimum number of exhaust inlets
- Unvented filling-time estimate (screening)

### Testing & Acceptance tab
- NFPA 92 Chapter 8 acceptance testing checklist
- Periodic ITM frequencies (dedicated semiannual / non-dedicated annual)
- Chapter 7 documentation requirements
- Typical atrium exhaust acceptance measurements

### Data checklist tab
Eleven categories covering the information typically required for a compliant design submission:

1. Project & regulatory data  
2. Geometry & space description  
3. Design fire scenario  
4. Smoke-layer design criteria  
5. Calculation methods & results  
6. System components & equipment  
7. Controls, detection & activation  
8. Testing, commissioning & documentation  
9. Special conditions & coordination  

Progress is tracked live in the session.

### Reference data tab
Key tabulated values: typical design-fire sizes, t² coefficients, tenability criteria, plume equations, makeup-air guidance.

### Excel report
Multi-sheet workbook: Cover, Inputs, Calculation summary, Data checklist status, Reference.

## Quick start

```bash
cd nfpa-92b-smoke-management
pip install -r requirements.txt
streamlit run app.py
```

Open the URL shown in the terminal (usually http://localhost:8501).

## Folder structure

```
nfpa-92b-smoke-management/
├── app.py                 # Streamlit application
├── requirements.txt
├── README.md
├── data/
│   └── reference_data.json
└── exports/               # Generated Excel reports (optional)
```

## Important limitations

- This is an **educational / preliminary-design aid**.
- Algebraic methods have geometric and fire-size limitations. Complex geometries, balcony-spill or window plumes, stratification, and full tenability analyses normally require zone models or CFD accepted by the authority having jurisdiction.
- Plugholing and filling-time relations are simplified screening forms; always verify against the exact equations in the NFPA 92 edition applicable to the project.
- Always obtain and apply the official NFPA standard and any local amendments.
- Final designs must be produced by competent professionals and approved by the AHJ.

## Disclaimer

The authors accept no liability for design decisions made using this tool.  
Final designs must be produced by competent professionals, verified as required, and approved by the relevant fire authority / building control body.
