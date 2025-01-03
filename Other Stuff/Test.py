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
kinase_group = input("What is the Kinase Group")
kinase_family = input("What is the Kinase Family")
kinase_name = input("What is the Kinase Name")
ligand_expo_id = input("What is the Ligand Expo ID")



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


## Case Study 2

# get the average interaction fingerprint
def average_n_interactions_per_residue(structure_klifs_ids):
    """
    Generate residue position x interaction type matrix that contains
    the average number of interactions per residue and interaction type.

    Parameters
    ----------
    structure_klifs_ids : list of int
        Structure KLIFS IDs.

    Returns
    -------
    pandas.DataFrame
        Average number of interactions per residue (rows) and interaction type (columns).
    """

    # Get IFP (is returned from KLIFS as string)
    ifps = session.interactions.by_structure_klifs_id(structure_klifs_ids)
    # Split string into list of int (0, 1): structures x IFP bits matrix
    ifps = pd.DataFrame(ifps["interaction.fingerprint"].apply(lambda x: list(x)).to_list())
    ifps = ifps.astype("int32")
    # Sum up all interaction per bit position and normalize by number of IFPs
    ifp_relative = (ifps.sum() / len(ifps)).to_list()
    # Transform aggregated IFP into residue position x interaction type matrix
    residue_feature_matrix = pd.DataFrame(
        [ifp_relative[i : i + 7] for i in range(0, len(ifp_relative), 7)], index=range(1, 86)
    )
    # Add interaction type names
    columns = session.interactions.interaction_types["interaction.name"].to_list()
    residue_feature_matrix.columns = columns
    return residue_feature_matrix

## Case Study 2
residue_feature_matrix = average_n_interactions_per_residue(structure_klifs_ids)
display(Markdown(
        f"Calculated average interactions per interaction type (feature) and residue position:\n\n"
        f"**Interaction types** (columns): {', '.join(residue_feature_matrix.columns)}\n\n"
        f"**Residue positions** (rows): {residue_feature_matrix.index}"
    )
)
# NBVAL_CHECK_OUTPUT
residue_feature_matrix

## My Table code for average interaction finger print
print(tabulate(residue_feature_matrix, ['Residue', 'Apolar Contact', 'Aromatic face-to-face', 'Aromatic edge-to-face',
                                        'Hydrogen bond donor (protein)', 'Hydrogen bond acceptor (protein)',
                                        'Protein cation - ligand anion', 'Protein anion - ligand cation'], 'github'))


my_df = pd.DataFrame(residue_feature_matrix)
my_df.to_csv("residues.csv", "w")
#text_file = open("Residues.xlsx", "w")
#text_file.write(residue_feature_matrix)
#text_file.close()

## Case Study 3

# trim structure list down to only ones bound to ligand of interest
subset = structures[(structures["ligand.expo_id"] == ligand_expo_id)]
subset = subset.sort_values(by="structure.resolution")

print(tabulate(subset, ['structure.klifs_id', 'structure.pdb_id',
               'structure.alternate_model', 'structure.chain', 'species.klifs', 'kinase.klifs_id',
               'kinase.klifs_name', 'kinase.names', 'kinase.family', 'kinase.group', 'structure.pocket',
                'ligand.expo_id', 'ligand_allosetric.expo_id', 'ligand.klifs_id', 'ligand_allosteric.klifs_id',
                'ligand.name', 'ligand_allosteric.name', 'structure.dfg', 'structure.ac_helix',
                'structure.resolution',	'structure.qualityscore','structure.missing_residues',
                'structure.missing_atoms','structure.rmsd1','structure.rmsd2','interaction.fingerprint',
                'structure.front','structure.gate','structure.back','structure.fp_i','structure.fp_ii',
                'structure.bp_i_a','structure.bp_i_b',	'structure.bp_ii_in',	'structure.bp_ii_a_in',
                'structure.bp_ii_b_in',	'structure.bp_ii_out',	'structure.bp_ii_b',	'structure.bp_iii',
                'structure.bp_iv',	'structure.bp_v',	'structure.grich_distance',	'structure.grich_angle',
                    'structure.grich_rotation',	'structure.filepath',	'structure.curation_flag']))

# return the ID of the best resolved ligand
structure_pdb_id = (
    structures[(structures["ligand.expo_id"] == ligand_expo_id)]
    .sort_values(by="structure.resolution")
    .iloc[0]["structure.pdb_id"]
)
message = f"Structure PDB ID for best resolved {ligand_expo_id}-bound {kinase_name}: {structure_pdb_id}"
display_markdown(message)


## Case Study 6

# Profiling data
bioactivities = session.bioactivities.by_ligand_expo_id(ligand_expo_id)
display(
    Markdown(
        f"Number of bioactivity values for {ligand_expo_id}: {len(bioactivities)}\n\n"
        f"Show example bioactivities:\n\n"
    )
)
bioactivities.sort_values("ligand.bioactivity_standard_value").head()


ACTIVITY_CUTOFF = 100
bioactivities_active = bioactivities[
    bioactivities["ligand.bioactivity_standard_value"] < ACTIVITY_CUTOFF
]
display(Markdown(("Number of measurements with high activity per kinase:")))
n_bioactivities_per_target = (
    bioactivities_active.groupby("kinase.pref_name").size().sort_values(ascending=True)
)
n_bioactivities_per_target

display_markdown(f"Off-targets of {ligand_expo_id} based on profiling data:")
bioactivities_active[
    bioactivities_active['kinase.pref_name'] != "Epidermal growth factor receptor erbB1"
].sort_values(["ligand.bioactivity_standard_value"])

print(tabulate(bioactivities_active, ['kinase.pref_name', 'kinase.uniprot', 'kinase.chembl_id', 'ligand.chembl_id',
               'ligand.bioactivity_standard_type', 'ligand.bioactivity_standard_relation',
               'ligand.bioactivity_standard_value', 'ligand.bioactivity_standard_units',
               'ligand.bioactivity_pchembl_value', 'species.chembl', 'ligand.expo_id'], "github"))




