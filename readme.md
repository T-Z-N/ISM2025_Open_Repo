# ISM2025 - Ontology Analysis and Taxonomy Processing

This repository contains tools and scripts for processing ontology data, particularly focused on taxonomy hierarchies and terminology annotations. The project includes SPARQL queries, Python processing scripts, and resulting data files for ontology analysis.

## Project Structure

```
├── AnnotationQuery1.rq              # SPARQL query for annotations
├── TaxonomyQuery2.rq                # SPARQL query for taxonomy relationships
├── ESPR_BFO_Classifications_Final.xlsx  # Final classification data
├── taxonomy_hierarchy.xlsx          # Processed taxonomy hierarchy
├── OntologyFiles/                   # Ontology data files
│   ├── WithoutTaxonomicalRelations.ttl
│   └── WithTaxonomicalRelations.ttl
├── PythonScripts/                   # Processing scripts
│   ├── csvToexcel_taxonomy.py       # Converts taxonomy CSV to Excel
│   └── csvToexcel_terminology.py    # Converts terminology CSV to Excel
└── Raw Query Results/               # Raw SPARQL query outputs
    ├── raw_class_annotations.xlsx
    ├── taxonomy_query_results.csv
    └── terminology_query_results.csv
```

## Features

### Ontology Files
The project works with two main ontology files in Turtle (TTL) format:

- **WithoutTaxonomicalRelations.ttl**: Base ontology file containing class definitions and annotations without explicit taxonomical relationships
- **WithTaxonomicalRelations.ttl**: Enhanced ontology file that includes taxonomical hierarchies and relationships between classes

These files serve as the data source for SPARQL queries and contain the semantic information processed by the analysis scripts.

### SPARQL Queries
- **AnnotationQuery1.rq**: Extracts annotation data from ontology files
- **TaxonomyQuery2.rq**: Retrieves taxonomical relationships

### Python Processing Scripts

#### csvToexcel_taxonomy.py
Converts taxonomy query results to a formatted Excel file with the following features:
- Reads from `taxonomy_query_results.csv`
- Outputs to `taxonomy_hierarchy.xlsx`
- Auto-adjusts column widths
- Merges cells for superclass columns with identical values

#### csvToexcel_terminology.py
Processes terminology annotations with enhanced formatting:
- Reads from `query-result.csv` 
- Outputs to `class_annotations.xlsx`
- Merges cells for class columns with identical values

## Usage

### Prerequisites
```bash
pip install pandas openpyxl
```

### Working with Ontology Files

The ontology files in the `OntologyFiles/` directory are in Turtle (TTL) format and can be:
- Loaded into SPARQL endpoints for querying
- Opened with ontology editors like Protégé
- Processed with RDF libraries in Python (e.g., rdflib)

### Running the Scripts

1. **Process Taxonomy Data:**
   ```bash
   cd PythonScripts
   python csvToexcel_taxonomy.py
   ```

2. **Process Terminology Data:**
   ```bash
   cd PythonScripts
   python csvToexcel_terminology.py
   ```

### Input Requirements
- Ensure CSV files are present in the same directory as the scripts
- For taxonomy processing: `taxonomy_query_results.csv`
- For terminology processing: `query-result.csv`

## Data Flow

1. **Ontology Files (TTL)** → Source data in RDF/Turtle format
2. **SPARQL Queries** → Extract data from TTL ontology files
3. **Raw CSV Results** → Store query outputs
4. **Python Scripts** → Process and format data
5. **Excel Output** → Final formatted files for analysis

## Output Files

- **taxonomy_hierarchy.xlsx**: Formatted taxonomy relationships with merged superclass cells
- **class_annotations.xlsx**: Processed terminology annotations with grouped classes
