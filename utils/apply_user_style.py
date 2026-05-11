"""Apply a user-supplied PPTX's theme (colors + fonts) to a base template.

The output is a copy of `base_template` with its `ppt/theme/theme1.xml` swapped
for the one in `user_pptx`. All other parts (slide masters, layouts,
placeholder names) are preserved so downstream layout-filler code keeps
working against the canonical template structure.
"""

from __future__ import annotations

from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED


THEME_PATH = "ppt/theme/theme1.xml"


def apply_user_style(
    user_pptx: str | Path,
    base_template: str | Path,
    output_path: str | Path,
) -> Path:
    user_pptx = Path(user_pptx)
    base_template = Path(base_template)
    output_path = Path(output_path)

    with ZipFile(user_pptx, "r") as z:
        if THEME_PATH not in z.namelist():
            raise ValueError(
                f"User template {user_pptx} does not contain {THEME_PATH}; "
                "cannot extract style."
            )
        user_theme = z.read(THEME_PATH)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with ZipFile(base_template, "r") as zin, ZipFile(
        output_path, "w", ZIP_DEFLATED
    ) as zout:
        for name in zin.namelist():
            if name == THEME_PATH:
                zout.writestr(name, user_theme)
            else:
                zout.writestr(name, zin.read(name))

    return output_path
