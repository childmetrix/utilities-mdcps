# utilities-mdcps

Mississippi (MDCPS) specific R utilities for ChildMetrix projects.

## Contents

### Functions (`/functions`)
- **functions_mdcps.R** - MDCPS-specific helper functions

### Templates (`/templates`)
- **r_script_template_mdcps.R** - MDCPS R script template
- **Letterhead - MDCPS - 20250608.dotx** - MDCPS letterhead template

## Usage

Source these files in your MDCPS project scripts:

```r
# Load core utilities first
source("D:/repo_childmetrix/utilities-core/loader.R")

# Then load MDCPS-specific utilities
source("D:/repo_childmetrix/utilities-mdcps/functions/functions_mdcps.R")
```

## Related Projects

MDCPS consent decree commitment projects (ms-1.3.a, ms-2.8.a, etc.)

## Maintenance

When adding MDCPS-specific functions:
1. Add to `functions/functions_mdcps.R`
2. Update this README
3. Test with relevant MDCPS projects before committing
