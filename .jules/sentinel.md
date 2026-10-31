# Sentinel's journal

Critical security learnings for this codebase only.

## 2026-10-31 - YAML aliases amplify at str(), not at parse time
**Vulnerability:** A label or site file of under 500 bytes could use nested
YAML anchors to make one field expand to hundreds of millions of characters,
exhausting memory in the CLI or GUI.
**Learning:** `yaml.safe_load` blocks code execution but not alias
amplification. It builds shared references cheaply, so parsing looked
harmless; the blow-up happened later, wherever a value was turned into text:
`Label.from_dict`, `Site.__post_init__`, `str(site_id)` and `{value!r}` in
error messages. Guarding each of those would always miss the next one.
**Prevention:** Read YAML only through `sites.load_yaml_document()`, which
measures the expanded size (memoised by identity, so the check itself cannot
blow up) and rejects documents out of proportion to their text. Any new YAML
reader must use it rather than calling `yaml.safe_load` directly.
