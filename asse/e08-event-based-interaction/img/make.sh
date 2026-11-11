#!/bin/sh
# Render the rule diagrams of the deck to SVG (needs plantuml).
# PlantUML measures text with the font "sans-serif" and fixes each label's width;
# the slides should show Gill Sans where it is installed, so swap in a font list.
set -e
cd "$(dirname "$0")"
plantuml -tsvg rules-*.puml
sed -i "s/font-family=\"sans-serif\"/font-family=\"'Gill Sans Nova','Gill Sans MT Pro','Gill Sans',sans-serif\"/g" rules-*.svg
