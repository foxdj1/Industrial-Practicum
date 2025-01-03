# Their imports
import logging
logging.getLogger("numexpr").setLevel(logging.ERROR)
from bravado.client import SwaggerClient
from IPython.display import display, Markdown, display_markdown
import pandas as pd
import nglview as nv
from rdkit import Chem
from rdkit.Chem.Draw import IPythonConsole, MolsToGridImage
import opencadd

# My imports
from tabulate import tabulate


# Setup for remote access
pd.set_option("display.max_columns", 50)

KLIFS_API_DEFINITIONS = "https://klifs.net/swagger/swagger.json"
KLIFS_CLIENT = SwaggerClient.from_url(KLIFS_API_DEFINITIONS, config={"validate_responses": False})

from opencadd.databases.klifs import setup_remote
session = setup_remote()

# Prompts to locate the correct kinase in the database
species = "Human"
kinase_group = 'TK'
kinase_family = 'EGFR'
kinase_name = 'EGFR'
ligand_expo_id = 'IRE'



## Tutorial Number 3

# Pull a list of all the kinases in the family then iterate through that list to find the one that we want
kinases = (
    KLIFS_CLIENT.Information.get_kinase_names(kinase_family=kinase_family, species=species)
    .response()
    .result
)
display(
    Markdown(
        f"Kinases in the {species.lower()} family {kinase_family} as a list of objects "
        f"that contain kinase-specific information:"
    )
)
kinases

# Use iterator to get first element of the list
kinase_klifs_id = next(kinase.kinase_ID for kinase in kinases if kinase.name == kinase_name)
display(Markdown(f"Kinase KLIFS ID for {kinase_name}: {kinase_klifs_id}"))
# NBVAL_CHECK_OUTPUT

#ligands = KLIFS_CLIENT.Information.get_ligands(kinase_family=kinase_family, species=species)
#ligand_expo_id = next(ligand.ligand_ID for ligand in ligands if ligand.name == ligand_name)

## Case Study 1

# Return the number of structures for a specific kinase
structures = session.structures.by_kinase_klifs_id(kinase_klifs_id)
display_markdown(f"Number of structures for the kinase {kinase_name}: {len(structures)}")


## Case Study 1

# Get the interaction fingerprint for the kinase-ligand structure
structure_klifs_ids = structures["structure.klifs_id"].to_list()
interaction_fingerprints = session.interactions.by_structure_klifs_id(structure_klifs_ids)
display_markdown(
        f"Number of IFPs for {kinase_name}: {len(interaction_fingerprints)}\n\n"
    )

interaction_fingerprints.head()
print(structures)
print(interaction_fingerprints)