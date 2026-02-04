# Configuration Files

This directory contains configuration files for the Omney Business Automation framework.

## Files

| File | Description |
|------|-------------|
| `config.json` | Main configuration file with all settings |
| `selectors.json` | UI element selectors for Playwright |
| `config.qa.json` | QA environment overrides |
| `config.uat.json` | UAT environment overrides |
| `config.prod.json` | Production environment overrides |

## config.json Structure

### environment
- `base_url`: Application URL
- `name`: Environment name (QA/UAT/PROD)

### browser
- `headless`: Run browser without UI (true/false)
- `slow_mo`: Delay between actions in milliseconds
- `args`: Browser launch arguments
- `no_viewport`: Use full browser window

### timeouts (in milliseconds)
- `element_visibility`: Wait for element to appear
- `page_navigation`: Wait for page navigation
- `dashboard_load`: Wait for dashboard to load
- `dropdown_open`: Wait for dropdown to open
- `short_delay`: Short wait (500ms)
- `medium_delay`: Medium wait (1000ms)
- `long_delay`: Long wait (2000ms)

### paths
- `reports_dir`: Directory for test reports
- `testcase_file`: Excel file with test cases
- `shared_invoice_data`: Shared data file between tests

### url_patterns
URL patterns for page navigation verification.

### defaults
Default values for test data.

### retry
- `max_attempts`: Number of retry attempts
- `dropdown_retry_delay`: Delay between dropdown retries

## selectors.json Structure

Organized by page/feature:
- `login`: Login page elements
- `navigation`: Navigation elements
- `invoice_form`: Invoice creation form
- `invoice_list`: Invoice listing table
- `invoice_details`: Invoice detail page
- `pay_invoice_page`: Pay invoice page
- `popups`: Modal/popup elements
- `data_extraction`: Labels for data extraction

## Usage

```python
import json
from pathlib import Path

# Load main config
config_path = Path(__file__).parent / "config" / "config.json"
with open(config_path) as f:
    config = json.load(f)

# Load selectors
selectors_path = Path(__file__).parent / "config" / "selectors.json"
with open(selectors_path) as f:
    selectors = json.load(f)

# Use in automation
base_url = config["environment"]["base_url"]
timeout = config["timeouts"]["page_navigation"]
login_selectors = selectors["login"]["email_input"]
```

## Environment Switching

To use a different environment, merge the environment-specific config:

```python
import json

def load_config(env="qa"):
    with open("config/config.json") as f:
        config = json.load(f)

    # Override with environment-specific settings
    env_file = f"config/config.{env}.json"
    if Path(env_file).exists():
        with open(env_file) as f:
            env_config = json.load(f)
        # Deep merge env_config into config
        for key, value in env_config.items():
            if isinstance(value, dict):
                config[key].update(value)
            else:
                config[key] = value

    return config

# Usage
config = load_config("uat")  # Use UAT environment
```

## Command Line Usage

```bash
# Run with QA config (default)
python omney_business_automation.py

# Run with UAT config
python omney_business_automation.py --env uat

# Run with PROD config
python omney_business_automation.py --env prod
```

## Adding New Selectors

When adding new selectors, follow this pattern:

```json
{
  "feature_name": {
    "element_name": [
      "primary_selector",
      "fallback_selector_1",
      "fallback_selector_2"
    ]
  }
}
```

Use arrays for elements that may have multiple valid selectors.

## Best Practices

1. **Never hardcode values** - Always use config files
2. **Use fallback selectors** - UI may change between releases
3. **Keep timeouts reasonable** - Too short = flaky tests, too long = slow tests
4. **Document changes** - Update this README when adding new configs
