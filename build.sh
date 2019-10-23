#!/bin/bash
DEFAULT_PPTX=$(find ~ -name default.pptx)
cp $DEFAULT_PPTX ./default.pptx
pyinstaller images2slides.spec
rm ./default.pptx