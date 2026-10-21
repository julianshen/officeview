# Slideshow capture checklist (D4)

PowerPoint reference frames for slideshow validation.

## Corpus decks

For each corpus deck under tests/fixtures/slideshow/, capture per slide and step at 0/25/50/75/100 % of each effect.

- tdf104788 (benchmark): 6 slides with notes, entrance presets 10/34/37/42/47/58
- tdf115394: fade with spd only and spd+p14:dur (250, 4250)
- tdf142915: advClick=0 advTm=1000, p14:dur 0 and 10
- tdf44223: split, p:sndAc, bldP animBg=1, entrance preset 2
- bnc591147: interactiveSeq onClick, mediacall cmd, video.avi, audio.mp3
- connector-shape-animations: preset 22 on connectors
- tdf104792-smart-art-animation: SmartArt build
- narration: p:audio, p14:media, media1.m4a, transitions
- airplane-transition: p15:prstTrans airplane
- tdf156808: p14:flythrough
- tdf149588_transparentSolidFill: p159:morph

## Behavioural questions

1. spd durations (fast/med/slow) vs p14:dur
2. advTm with pending clicks
3. trigger vs hyperlink precedence
4. kiosk restart
5. hyperlinks and triggers in kiosk mode
6. Previous into a slide
7. spd="med" beside p14:dur="700"
EOF
echo done