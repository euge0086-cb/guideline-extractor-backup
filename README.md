# 📚 Guideline Reference Extractor

App web para extraer, enriquecer y clasificar referencias de guías clínicas en PDF.

## Qué hace

1. **Extrae** todas las referencias de un PDF de guía clínica (ESC, ACC/AHA u otras)
2. **Enriquece** con metadatos: PMID, DOI, autores, año, revista (vía PubMed + CrossRef)
3. **Clasifica** automáticamente: ECA primario, ECA secundario, meta-análisis, registro, guía
4. **Exporta** un Excel con 4 hojas: base completa, ECAs primarios, resumen, instrucciones


## Archivos del repositorio

```
├── app.py              # Interfaz Streamlit
├── pipeline.py         # Motor de extracción y enriquecimiento
├── requirements.txt    # Dependencias Python
└── README.md           # Este archivo
```

## Uso local (opcional)

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Notas

- Sin API key de NCBI el límite es 3 req/s (automáticamente respetado)
- Para guías grandes (>300 referencias) el procesamiento puede tardar 15-30 min
- La clasificación automática tiene ~80-90% de precisión; la columna "Tipo (manual)" permite correcciones

## Basado en

Mas-Llado C et al. *Representativeness in randomised clinical trials supporting acute coronary syndrome guidelines.* Eur Heart J Qual Care Clin Outcomes. 2023;9:796-805. https://doi.org/10.1093/ehjqcco/qcad007
