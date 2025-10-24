# utilities-mdcps

Mississippi (MDCPS) specific R utilities for ChildMetrix projects.

**Repository**: [github.com/childmetrix/utilities-mdcps](https://github.com/childmetrix/utilities-mdcps)

## Purpose

This repo contains Mississippi Department of Child Protection Services (MDCPS) specific functions, templates, and resources used across all Mississippi consent decree commitment projects.

## Contents

### Functions Directory

**functions_mdcps.R**
- MDCPS-specific data processing functions
- Mississippi data structure assumptions
- MDCPS reporting utilities
- State-specific calculations and transformations
- Use: `source("D:/repo_childmetrix/utilities-mdcps/functions/functions_mdcps.R")`

### Templates Directory

**r_script_template_mdcps.R**
- MDCPS-specific R script template
- Includes MDCPS-specific setup and sourcing
- Standard sections for MDCPS analysis workflow

**Letterhead - MDCPS - 20250608.dotx**
- Official MDCPS letterhead template for Word documents
- Use for formal reports and correspondence

## Usage in MDCPS Projects

All Mississippi projects should source both core and MDCPS utilities:

```r
# Load core utilities (always first)
source("D:/repo_childmetrix/utilities-core/loader.R")
source("D:/repo_childmetrix/utilities-core/functions/generic_functions.R")

# Load MDCPS-specific utilities
source("D:/repo_childmetrix/utilities-mdcps/functions/functions_mdcps.R")
```

## Related Projects

MDCPS consent decree commitment projects following naming pattern: `ms-{commitment-number}`

Examples:
- `ms-1.3a` - Commitment 1.3.a
- `ms-2.8a` - Commitment 2.8.a
- etc.

These projects are typically stored in: `D:\repo_mdcps_suspension_period\`

## When to Add Functions Here

Add functions to utilities-mdcps when they:
- Are specific to Mississippi/MDCPS data or requirements
- Will be reused across multiple MDCPS commitment projects
- Contain MDCPS-specific business logic or calculations
- Are stable and well-tested

## When NOT to Add Functions Here

Do NOT add functions here if they:
- Are generic and could be used by other states → use utilities-core
- Are specific to a single commitment project → keep in that project repo
- Are experimental or untested → test in project first
- Contain confidential data or logic → keep in project repo

## Maintenance

When modifying MDCPS utilities:
1. Test changes across multiple MDCPS projects before committing
2. Ensure changes don't break existing commitment analyses
3. Update this CLAUDE.md if you add new files or significant functions
4. Consider impact on historical analyses and reports
