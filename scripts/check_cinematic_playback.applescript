tell application "Microsoft PowerPoint"
 set deck to presentation "AGROVECTOR_CINEMATIC.pptx"
 set ss to slide show settings of deck
 set show type of ss to slide show type window
 set starting slide of ss to 1
 set ending slide of ss to 22
 set win to run slide show ss
 delay 2
 return properties of slide show view of win
end tell
