tell application "Microsoft PowerPoint"
 try
  close presentation "AGROVECTOR_CINEMATIC.pptx" saving no
 end try
 open POSIX file "/Users/daniilvazdaev/Desktop/уник/Бизнес-планирование/AgroVector/presentations/AGROVECTOR_CINEMATIC.pptx"
 set deck to active presentation
 repeat with i from 1 to 22
  set settings to animation play settings of animation settings of shape 1 of slide i of deck
  set play on entry of settings to true
  set loop until stopped of settings to true
 end repeat
 save deck
 return {name of deck, count of slides of deck, play on entry of animation play settings of animation settings of shape 1 of slide 1 of deck, loop until stopped of animation play settings of animation settings of shape 1 of slide 1 of deck}
end tell
