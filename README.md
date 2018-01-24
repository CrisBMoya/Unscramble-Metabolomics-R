# Unscramble Metabolomics with R

**This script has been redacted specifically for association of phenotypes to metabolites using PLS-DA regression.**

Other uses should be plausible.

This script mimics what The Unscrambler X do. The calculations and output plots expected.

**Instructions**: the Excel should have a specific structure:

- In this file the Rows represent each Sample or Replicate.
- The First Column should have a name or identifier for every Sample or Replicate at the Column Name must be 'Name'
- Starting from the Second Column you should put all the information to be used as Predictors/Features, commonly the phenotypic traits.
- After the last Column with Predictors/Responses you should put all the information regarding Responses, commonly the metabolites -measured.
- **NOTE**: the predictors should ALWAYS be Categorical Data, meaning groups, quantiles, yes/no, presence/absence, high/mid/low, etc.

Data are stored in the selected output folder as sub folder representing every iteration process.
The iterations remove data that falls outside the confidence ellipse calculated.
The script re-do the calculations without this data, and filter again if needed.

**IGNORE** errors related to ggplot (ej: removed x rows containing missing values (geom_xxxx))
This are plotting errors and there are other plots who don't look pretty but shows the total data
