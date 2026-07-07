# Instructions

- Following Playwright test failed.
- Explain why, be concise, respect Playwright best practices.
- Provide a snippet of code with the fix, if possible.

# Test info

- Name: behaviors.spec.ts >> headers and footers >> single click is gated; double-click enters with chrome; body double-click exits
- Location: e2e/behaviors.spec.ts:226:7

# Error details

```
Error: expect(received).toBe(expected) // Object.is equality

Expected: false
Received: true
```

# Page snapshot

```yaml
- generic [ref=e3]:
  - banner [ref=e4]:
    - strong [ref=e5]: DocxInWeb
    - button "Choose File" [ref=e6]
    - generic [ref=e7]:
      - text: Zoom
      - combobox "Zoom" [ref=e8]:
        - option "50%"
        - option "75%"
        - option "100%" [selected]
        - option "125%"
        - option "150%"
    - generic [ref=e9]:
      - checkbox "Edit" [checked] [ref=e10]
      - text: Edit
    - generic "Show review comments" [ref=e11]:
      - checkbox "Comments" [checked] [ref=e12]
      - text: Comments
    - generic "Show tracked changes (insertions underlined, deletions struck)" [ref=e13]:
      - checkbox "Tracked changes" [ref=e14]
      - text: Tracked changes
    - button "Print" [ref=e15]
    - generic [ref=e16]: 4 pages
  - generic [ref=e17]:
    - generic [ref=e18]:
      - button "home" [ref=e19] [cursor=pointer]
      - button "insert" [ref=e20] [cursor=pointer]
      - button "layout" [ref=e21] [cursor=pointer]
    - button "↶" [ref=e23] [cursor=pointer]
    - button "↷" [ref=e24] [cursor=pointer]
    - combobox "Paragraph style" [ref=e26] [cursor=pointer]:
      - option "Normal"
      - option "Heading 1" [selected]
      - option "Heading 2"
      - option "Heading 3"
      - option "Heading 4"
      - option "Heading 5"
      - option "Heading 6"
      - option "Title"
    - combobox "Font" [ref=e27] [cursor=pointer]:
      - option "Font" [disabled] [selected]
      - option "Arial"
      - option "Calibri"
      - option "Cambria"
      - option "Courier New"
      - option "Garamond"
      - option "Georgia"
      - option "Helvetica"
      - option "Times New Roman"
      - option "Trebuchet MS"
      - option "Verdana"
    - combobox "Font size" [ref=e28] [cursor=pointer]:
      - option "Size" [disabled] [selected]
      - option "8"
      - option "9"
      - option "10"
      - option "10.5"
      - option "11"
      - option "12"
      - option "14"
      - option "16"
      - option "18"
      - option "20"
      - option "24"
      - option "28"
      - option "36"
      - option "48"
    - button "B" [ref=e30] [cursor=pointer]
    - button "I" [ref=e31] [cursor=pointer]
    - button "U" [ref=e32] [cursor=pointer]
    - button "S" [ref=e33] [cursor=pointer]
    - button "x2" [ref=e34] [cursor=pointer]:
      - generic [ref=e35]:
        - text: x
        - superscript [ref=e36]: "2"
    - button "x2" [ref=e37] [cursor=pointer]:
      - generic [ref=e38]:
        - text: x
        - subscript [ref=e39]: "2"
    - button "Clear formatting" [ref=e40] [cursor=pointer]:
      - img [ref=e41]
    - combobox "Change case" [ref=e44] [cursor=pointer]:
      - option "Aa" [disabled] [selected]
      - option "UPPERCASE"
      - option "lowercase"
      - option "Title Case"
    - generic "Text color" [ref=e45] [cursor=pointer]:
      - generic [ref=e46]: A
      - textbox "A": "#000000"
    - button "Highlight color" [ref=e48] [cursor=pointer]:
      - img [ref=e49]
    - button "≡" [ref=e53] [cursor=pointer]
    - button "≣" [ref=e54] [cursor=pointer]
    - button "≢" [ref=e55] [cursor=pointer]
    - button "☰" [ref=e56] [cursor=pointer]
    - button "Decrease indent" [ref=e58] [cursor=pointer]:
      - img [ref=e59]
    - button "Increase indent" [ref=e62] [cursor=pointer]:
      - img [ref=e63]
    - combobox "Line & paragraph spacing" [ref=e66] [cursor=pointer]:
      - option "↕" [disabled] [selected]
      - option "Single"
      - option "1.15"
      - option "1.5"
      - option "Double"
      - option "Add space before"
      - option "Remove space before"
      - option "Add space after"
      - option "Remove space after"
    - button "Bulleted list" [ref=e67] [cursor=pointer]:
      - img [ref=e68]
    - button "Numbered list" [ref=e73] [cursor=pointer]:
      - img [ref=e74]:
        - generic [ref=e75]: "1"
        - generic [ref=e76]: "2"
        - generic [ref=e77]: "3"
    - button "Download" [ref=e80] [cursor=pointer]
  - generic [ref=e83]:
    - textbox [active] [ref=e84]
    - generic [ref=e85]:
      - generic [ref=e87]:
        - generic [ref=e88]: DocxInWeb
        - generic [ref=e89]: Rendering
        - generic [ref=e90]: Test
        - generic [ref=e91]: This
        - generic [ref=e92]: document
        - generic [ref=e93]: exercises
        - generic [ref=e94]: the
        - generic [ref=e95]: fidelity-critical
        - generic [ref=e96]: "features:"
        - generic [ref=e97]: bold
        - generic [ref=e98]: ","
        - generic [ref=e99]: italic
        - generic [ref=e100]: ","
        - generic [ref=e101]: underline
        - generic [ref=e102]: ","
        - generic [ref=e103]: colored
        - generic [ref=e104]: text
        - generic [ref=e105]: ","
        - generic [ref=e107]: highlight
        - generic [ref=e108]: ","
        - generic [ref=e109]: and
        - generic [ref=e110]: superscript
        - generic [ref=e111]: .
        - generic [ref=e112]: Justified
        - generic [ref=e113]: paragraph
        - generic [ref=e114]: Lorem
        - generic [ref=e115]: ipsum
        - generic [ref=e116]: dolor
        - generic [ref=e117]: sit
        - generic [ref=e118]: amet,
        - generic [ref=e119]: consectetur
        - generic [ref=e120]: adipiscing
        - generic [ref=e121]: elit,
        - generic [ref=e122]: sed
        - generic [ref=e123]: do
        - generic [ref=e124]: eiusmod
        - generic [ref=e125]: tempor
        - generic [ref=e126]: incididunt
        - generic [ref=e127]: ut
        - generic [ref=e128]: labore
        - generic [ref=e129]: et
        - generic [ref=e130]: dolore
        - generic [ref=e131]: magna
        - generic [ref=e132]: aliqua.
        - generic [ref=e133]: Ut
        - generic [ref=e134]: enim
        - generic [ref=e135]: ad
        - generic [ref=e136]: minim
        - generic [ref=e137]: veniam,
        - generic [ref=e138]: quis
        - generic [ref=e139]: nostrud
        - generic [ref=e140]: exercitation
        - generic [ref=e141]: ullamco
        - generic [ref=e142]: laboris
        - generic [ref=e143]: nisi
        - generic [ref=e144]: ut
        - generic [ref=e145]: aliquip
        - generic [ref=e146]: ex
        - generic [ref=e147]: ea
        - generic [ref=e148]: commodo
        - generic [ref=e149]: consequat.
        - generic [ref=e150]: Duis
        - generic [ref=e151]: aute
        - generic [ref=e152]: irure
        - generic [ref=e153]: dolor
        - generic [ref=e154]: in
        - generic [ref=e155]: reprehenderit
        - generic [ref=e156]: in
        - generic [ref=e157]: voluptate
        - generic [ref=e158]: velit
        - generic [ref=e159]: esse
        - generic [ref=e160]: cillum
        - generic [ref=e161]: dolore
        - generic [ref=e162]: eu
        - generic [ref=e163]: fugiat
        - generic [ref=e164]: nulla
        - generic [ref=e165]: pariatur.
        - generic [ref=e166]: Excepteur
        - generic [ref=e167]: sint
        - generic [ref=e168]: occaecat
        - generic [ref=e169]: cupidatat
        - generic [ref=e170]: non
        - generic [ref=e171]: proident,
        - generic [ref=e172]: sunt
        - generic [ref=e173]: in
        - generic [ref=e174]: culpa
        - generic [ref=e175]: qui
        - generic [ref=e176]: officia
        - generic [ref=e177]: deserunt
        - generic [ref=e178]: mollit
        - generic [ref=e179]: anim
        - generic [ref=e180]: id
        - generic [ref=e181]: est
        - generic [ref=e182]: laborum.
        - generic [ref=e183]: Lists
        - generic [ref=e184]: "1."
        - generic [ref=e185]: First
        - generic [ref=e186]: numbered
        - generic [ref=e187]: item
        - generic [ref=e188]: "2."
        - generic [ref=e189]: Second
        - generic [ref=e190]: numbered
        - generic [ref=e191]: item
        - generic [ref=e192]: a)
        - generic [ref=e193]: Nested
        - generic [ref=e194]: letter
        - generic [ref=e195]: item
        - generic [ref=e196]: b)
        - generic [ref=e197]: Another
        - generic [ref=e198]: nested
        - generic [ref=e199]: item
        - generic [ref=e200]: "3."
        - generic [ref=e201]: Back
        - generic [ref=e202]: to
        - generic [ref=e203]: top
        - generic [ref=e204]: level
        - generic [ref=e205]: ●
        - generic [ref=e206]: Bullet
        - generic [ref=e207]: one
        - generic [ref=e208]: ●
        - generic [ref=e209]: Bullet
        - generic [ref=e210]: two
        - generic [ref=e211]: Divider
        - generic [ref=e212]: line
        - generic [ref=e213]: below
        - generic [ref=e214]: (paragraph
        - generic [ref=e215]: bottom
        - generic [ref=e216]: "border):"
        - generic [ref=e218]: Table
        - generic [ref=e220]: Feature
        - generic [ref=e226]: Status
        - generic [ref=e232]: Notes
        - generic [ref=e238]: Pagination
        - generic [ref=e243]: Working
        - generic [ref=e248]: Real
        - generic [ref=e249]: page
        - generic [ref=e250]: boxes
        - generic [ref=e251]: with
        - generic [ref=e252]: measured
        - generic [ref=e253]: line
        - generic [ref=e254]: breaking
        - generic [ref=e260]: Page
        - generic [ref=e261]: numbers
        - generic [ref=e266]: Working
        - generic [ref=e271]: PAGE
        - generic [ref=e272]: /
        - generic [ref=e273]: NUMPAGES
        - generic [ref=e274]: fields
        - generic [ref=e275]: resolved
        - generic [ref=e276]: at
        - generic [ref=e277]: layout
        - generic [ref=e278]: time
        - generic [ref=e287]: DocxInWeb
        - generic [ref=e288]: Fidelity
        - generic [ref=e289]: Sample
        - generic [ref=e290]: Page
        - generic [ref=e291]: "1"
        - generic [ref=e292]: of
        - generic [ref=e293]: "4"
      - generic [ref=e295]:
        - generic [ref=e296]: Page
        - generic [ref=e297]: "2"
        - generic [ref=e298]: This
        - generic [ref=e299]: paragraph
        - generic [ref=e300]: starts
        - generic [ref=e301]: page
        - generic [ref=e302]: two
        - generic [ref=e303]: after
        - generic [ref=e304]: an
        - generic [ref=e305]: explicit
        - generic [ref=e306]: page
        - generic [ref=e307]: break.
        - generic [ref=e308]: The
        - generic [ref=e309]: footer
        - generic [ref=e310]: below
        - generic [ref=e311]: should
        - generic [ref=e312]: read
        - generic [ref=e313]: “Page
        - generic [ref=e314]: "2"
        - generic [ref=e315]: of
        - generic [ref=e316]: N”.
        - generic [ref=e317]: Filler
        - generic [ref=e318]: paragraph
        - generic [ref=e319]: "1"
        - generic [ref=e320]: —
        - generic [ref=e321]: long
        - generic [ref=e322]: enough
        - generic [ref=e323]: content
        - generic [ref=e324]: to
        - generic [ref=e325]: force
        - generic [ref=e326]: natural
        - generic [ref=e327]: pagination
        - generic [ref=e328]: across
        - generic [ref=e329]: multiple
        - generic [ref=e330]: pages
        - generic [ref=e331]: so
        - generic [ref=e332]: that
        - generic [ref=e333]: widow
        - generic [ref=e334]: control,
        - generic [ref=e335]: page
        - generic [ref=e336]: fill
        - generic [ref=e337]: and
        - generic [ref=e338]: footers
        - generic [ref=e339]: can
        - generic [ref=e340]: be
        - generic [ref=e341]: verified
        - generic [ref=e342]: visually.
        - generic [ref=e343]: Filler
        - generic [ref=e344]: paragraph
        - generic [ref=e345]: "2"
        - generic [ref=e346]: —
        - generic [ref=e347]: long
        - generic [ref=e348]: enough
        - generic [ref=e349]: content
        - generic [ref=e350]: to
        - generic [ref=e351]: force
        - generic [ref=e352]: natural
        - generic [ref=e353]: pagination
        - generic [ref=e354]: across
        - generic [ref=e355]: multiple
        - generic [ref=e356]: pages
        - generic [ref=e357]: so
        - generic [ref=e358]: that
        - generic [ref=e359]: widow
        - generic [ref=e360]: control,
        - generic [ref=e361]: page
        - generic [ref=e362]: fill
        - generic [ref=e363]: and
        - generic [ref=e364]: footers
        - generic [ref=e365]: can
        - generic [ref=e366]: be
        - generic [ref=e367]: verified
        - generic [ref=e368]: visually.
        - generic [ref=e369]: Filler
        - generic [ref=e370]: paragraph
        - generic [ref=e371]: "3"
        - generic [ref=e372]: —
        - generic [ref=e373]: long
        - generic [ref=e374]: enough
        - generic [ref=e375]: content
        - generic [ref=e376]: to
        - generic [ref=e377]: force
        - generic [ref=e378]: natural
        - generic [ref=e379]: pagination
        - generic [ref=e380]: across
        - generic [ref=e381]: multiple
        - generic [ref=e382]: pages
        - generic [ref=e383]: so
        - generic [ref=e384]: that
        - generic [ref=e385]: widow
        - generic [ref=e386]: control,
        - generic [ref=e387]: page
        - generic [ref=e388]: fill
        - generic [ref=e389]: and
        - generic [ref=e390]: footers
        - generic [ref=e391]: can
        - generic [ref=e392]: be
        - generic [ref=e393]: verified
        - generic [ref=e394]: visually.
        - generic [ref=e395]: Filler
        - generic [ref=e396]: paragraph
        - generic [ref=e397]: "4"
        - generic [ref=e398]: —
        - generic [ref=e399]: long
        - generic [ref=e400]: enough
        - generic [ref=e401]: content
        - generic [ref=e402]: to
        - generic [ref=e403]: force
        - generic [ref=e404]: natural
        - generic [ref=e405]: pagination
        - generic [ref=e406]: across
        - generic [ref=e407]: multiple
        - generic [ref=e408]: pages
        - generic [ref=e409]: so
        - generic [ref=e410]: that
        - generic [ref=e411]: widow
        - generic [ref=e412]: control,
        - generic [ref=e413]: page
        - generic [ref=e414]: fill
        - generic [ref=e415]: and
        - generic [ref=e416]: footers
        - generic [ref=e417]: can
        - generic [ref=e418]: be
        - generic [ref=e419]: verified
        - generic [ref=e420]: visually.
        - generic [ref=e421]: Filler
        - generic [ref=e422]: paragraph
        - generic [ref=e423]: "5"
        - generic [ref=e424]: —
        - generic [ref=e425]: long
        - generic [ref=e426]: enough
        - generic [ref=e427]: content
        - generic [ref=e428]: to
        - generic [ref=e429]: force
        - generic [ref=e430]: natural
        - generic [ref=e431]: pagination
        - generic [ref=e432]: across
        - generic [ref=e433]: multiple
        - generic [ref=e434]: pages
        - generic [ref=e435]: so
        - generic [ref=e436]: that
        - generic [ref=e437]: widow
        - generic [ref=e438]: control,
        - generic [ref=e439]: page
        - generic [ref=e440]: fill
        - generic [ref=e441]: and
        - generic [ref=e442]: footers
        - generic [ref=e443]: can
        - generic [ref=e444]: be
        - generic [ref=e445]: verified
        - generic [ref=e446]: visually.
        - generic [ref=e447]: Filler
        - generic [ref=e448]: paragraph
        - generic [ref=e449]: "6"
        - generic [ref=e450]: —
        - generic [ref=e451]: long
        - generic [ref=e452]: enough
        - generic [ref=e453]: content
        - generic [ref=e454]: to
        - generic [ref=e455]: force
        - generic [ref=e456]: natural
        - generic [ref=e457]: pagination
        - generic [ref=e458]: across
        - generic [ref=e459]: multiple
        - generic [ref=e460]: pages
        - generic [ref=e461]: so
        - generic [ref=e462]: that
        - generic [ref=e463]: widow
        - generic [ref=e464]: control,
        - generic [ref=e465]: page
        - generic [ref=e466]: fill
        - generic [ref=e467]: and
        - generic [ref=e468]: footers
        - generic [ref=e469]: can
        - generic [ref=e470]: be
        - generic [ref=e471]: verified
        - generic [ref=e472]: visually.
        - generic [ref=e473]: Filler
        - generic [ref=e474]: paragraph
        - generic [ref=e475]: "7"
        - generic [ref=e476]: —
        - generic [ref=e477]: long
        - generic [ref=e478]: enough
        - generic [ref=e479]: content
        - generic [ref=e480]: to
        - generic [ref=e481]: force
        - generic [ref=e482]: natural
        - generic [ref=e483]: pagination
        - generic [ref=e484]: across
        - generic [ref=e485]: multiple
        - generic [ref=e486]: pages
        - generic [ref=e487]: so
        - generic [ref=e488]: that
        - generic [ref=e489]: widow
        - generic [ref=e490]: control,
        - generic [ref=e491]: page
        - generic [ref=e492]: fill
        - generic [ref=e493]: and
        - generic [ref=e494]: footers
        - generic [ref=e495]: can
        - generic [ref=e496]: be
        - generic [ref=e497]: verified
        - generic [ref=e498]: visually.
        - generic [ref=e499]: Filler
        - generic [ref=e500]: paragraph
        - generic [ref=e501]: "8"
        - generic [ref=e502]: —
        - generic [ref=e503]: long
        - generic [ref=e504]: enough
        - generic [ref=e505]: content
        - generic [ref=e506]: to
        - generic [ref=e507]: force
        - generic [ref=e508]: natural
        - generic [ref=e509]: pagination
        - generic [ref=e510]: across
        - generic [ref=e511]: multiple
        - generic [ref=e512]: pages
        - generic [ref=e513]: so
        - generic [ref=e514]: that
        - generic [ref=e515]: widow
        - generic [ref=e516]: control,
        - generic [ref=e517]: page
        - generic [ref=e518]: fill
        - generic [ref=e519]: and
        - generic [ref=e520]: footers
        - generic [ref=e521]: can
        - generic [ref=e522]: be
        - generic [ref=e523]: verified
        - generic [ref=e524]: visually.
        - generic [ref=e525]: Filler
        - generic [ref=e526]: paragraph
        - generic [ref=e527]: "9"
        - generic [ref=e528]: —
        - generic [ref=e529]: long
        - generic [ref=e530]: enough
        - generic [ref=e531]: content
        - generic [ref=e532]: to
        - generic [ref=e533]: force
        - generic [ref=e534]: natural
        - generic [ref=e535]: pagination
        - generic [ref=e536]: across
        - generic [ref=e537]: multiple
        - generic [ref=e538]: pages
        - generic [ref=e539]: so
        - generic [ref=e540]: that
        - generic [ref=e541]: widow
        - generic [ref=e542]: control,
        - generic [ref=e543]: page
        - generic [ref=e544]: fill
        - generic [ref=e545]: and
        - generic [ref=e546]: footers
        - generic [ref=e547]: can
        - generic [ref=e548]: be
        - generic [ref=e549]: verified
        - generic [ref=e550]: visually.
        - generic [ref=e551]: Filler
        - generic [ref=e552]: paragraph
        - generic [ref=e553]: "10"
        - generic [ref=e554]: —
        - generic [ref=e555]: long
        - generic [ref=e556]: enough
        - generic [ref=e557]: content
        - generic [ref=e558]: to
        - generic [ref=e559]: force
        - generic [ref=e560]: natural
        - generic [ref=e561]: pagination
        - generic [ref=e562]: across
        - generic [ref=e563]: multiple
        - generic [ref=e564]: pages
        - generic [ref=e565]: so
        - generic [ref=e566]: that
        - generic [ref=e567]: widow
        - generic [ref=e568]: control,
        - generic [ref=e569]: page
        - generic [ref=e570]: fill
        - generic [ref=e571]: and
        - generic [ref=e572]: footers
        - generic [ref=e573]: can
        - generic [ref=e574]: be
        - generic [ref=e575]: verified
        - generic [ref=e576]: visually.
        - generic [ref=e577]: Filler
        - generic [ref=e578]: paragraph
        - generic [ref=e579]: "11"
        - generic [ref=e580]: —
        - generic [ref=e581]: long
        - generic [ref=e582]: enough
        - generic [ref=e583]: content
        - generic [ref=e584]: to
        - generic [ref=e585]: force
        - generic [ref=e586]: natural
        - generic [ref=e587]: pagination
        - generic [ref=e588]: across
        - generic [ref=e589]: multiple
        - generic [ref=e590]: pages
        - generic [ref=e591]: so
        - generic [ref=e592]: that
        - generic [ref=e593]: widow
        - generic [ref=e594]: control,
        - generic [ref=e595]: page
        - generic [ref=e596]: fill
        - generic [ref=e597]: and
        - generic [ref=e598]: footers
        - generic [ref=e599]: can
        - generic [ref=e600]: be
        - generic [ref=e601]: verified
        - generic [ref=e602]: visually.
        - generic [ref=e603]: Filler
        - generic [ref=e604]: paragraph
        - generic [ref=e605]: "12"
        - generic [ref=e606]: —
        - generic [ref=e607]: long
        - generic [ref=e608]: enough
        - generic [ref=e609]: content
        - generic [ref=e610]: to
        - generic [ref=e611]: force
        - generic [ref=e612]: natural
        - generic [ref=e613]: pagination
        - generic [ref=e614]: across
        - generic [ref=e615]: multiple
        - generic [ref=e616]: pages
        - generic [ref=e617]: so
        - generic [ref=e618]: that
        - generic [ref=e619]: widow
        - generic [ref=e620]: control,
        - generic [ref=e621]: page
        - generic [ref=e622]: fill
        - generic [ref=e623]: and
        - generic [ref=e624]: footers
        - generic [ref=e625]: can
        - generic [ref=e626]: be
        - generic [ref=e627]: verified
        - generic [ref=e628]: visually.
        - generic [ref=e629]: Filler
        - generic [ref=e630]: paragraph
        - generic [ref=e631]: "13"
        - generic [ref=e632]: —
        - generic [ref=e633]: long
        - generic [ref=e634]: enough
        - generic [ref=e635]: content
        - generic [ref=e636]: to
        - generic [ref=e637]: force
        - generic [ref=e638]: natural
        - generic [ref=e639]: pagination
        - generic [ref=e640]: across
        - generic [ref=e641]: multiple
        - generic [ref=e642]: pages
        - generic [ref=e643]: so
        - generic [ref=e644]: that
        - generic [ref=e645]: widow
        - generic [ref=e646]: control,
        - generic [ref=e647]: page
        - generic [ref=e648]: fill
        - generic [ref=e649]: and
        - generic [ref=e650]: footers
        - generic [ref=e651]: can
        - generic [ref=e652]: be
        - generic [ref=e653]: verified
        - generic [ref=e654]: visually.
        - generic [ref=e655]: Filler
        - generic [ref=e656]: paragraph
        - generic [ref=e657]: "14"
        - generic [ref=e658]: —
        - generic [ref=e659]: long
        - generic [ref=e660]: enough
        - generic [ref=e661]: content
        - generic [ref=e662]: to
        - generic [ref=e663]: force
        - generic [ref=e664]: natural
        - generic [ref=e665]: pagination
        - generic [ref=e666]: across
        - generic [ref=e667]: multiple
        - generic [ref=e668]: pages
        - generic [ref=e669]: so
        - generic [ref=e670]: that
        - generic [ref=e671]: widow
        - generic [ref=e672]: control,
        - generic [ref=e673]: page
        - generic [ref=e674]: fill
        - generic [ref=e675]: and
        - generic [ref=e676]: footers
        - generic [ref=e677]: can
        - generic [ref=e678]: be
        - generic [ref=e679]: verified
        - generic [ref=e680]: visually.
        - generic [ref=e681]: Filler
        - generic [ref=e682]: paragraph
        - generic [ref=e683]: "15"
        - generic [ref=e684]: —
        - generic [ref=e685]: long
        - generic [ref=e686]: enough
        - generic [ref=e687]: content
        - generic [ref=e688]: to
        - generic [ref=e689]: force
        - generic [ref=e690]: natural
        - generic [ref=e691]: pagination
        - generic [ref=e692]: across
        - generic [ref=e693]: multiple
        - generic [ref=e694]: pages
        - generic [ref=e695]: so
        - generic [ref=e696]: that
        - generic [ref=e697]: widow
        - generic [ref=e698]: control,
        - generic [ref=e699]: page
        - generic [ref=e700]: fill
        - generic [ref=e701]: and
        - generic [ref=e702]: footers
        - generic [ref=e703]: can
        - generic [ref=e704]: be
        - generic [ref=e705]: verified
        - generic [ref=e706]: visually.
        - generic [ref=e707]: Filler
        - generic [ref=e708]: paragraph
        - generic [ref=e709]: "16"
        - generic [ref=e710]: —
        - generic [ref=e711]: long
        - generic [ref=e712]: enough
        - generic [ref=e713]: content
        - generic [ref=e714]: to
        - generic [ref=e715]: force
        - generic [ref=e716]: natural
        - generic [ref=e717]: pagination
        - generic [ref=e718]: across
        - generic [ref=e719]: multiple
        - generic [ref=e720]: pages
        - generic [ref=e721]: so
        - generic [ref=e722]: that
        - generic [ref=e723]: widow
        - generic [ref=e724]: control,
        - generic [ref=e725]: page
        - generic [ref=e726]: fill
        - generic [ref=e727]: and
        - generic [ref=e728]: footers
        - generic [ref=e729]: can
        - generic [ref=e730]: be
        - generic [ref=e731]: verified
        - generic [ref=e732]: visually.
        - generic [ref=e733]: Filler
        - generic [ref=e734]: paragraph
        - generic [ref=e735]: "17"
        - generic [ref=e736]: —
        - generic [ref=e737]: long
        - generic [ref=e738]: enough
        - generic [ref=e739]: content
        - generic [ref=e740]: to
        - generic [ref=e741]: force
        - generic [ref=e742]: natural
        - generic [ref=e743]: pagination
        - generic [ref=e744]: across
        - generic [ref=e745]: multiple
        - generic [ref=e746]: pages
        - generic [ref=e747]: so
        - generic [ref=e748]: that
        - generic [ref=e749]: widow
        - generic [ref=e750]: control,
        - generic [ref=e751]: page
        - generic [ref=e752]: fill
        - generic [ref=e753]: and
        - generic [ref=e754]: footers
        - generic [ref=e755]: can
        - generic [ref=e756]: be
        - generic [ref=e757]: verified
        - generic [ref=e758]: visually.
        - generic [ref=e759]: Filler
        - generic [ref=e760]: paragraph
        - generic [ref=e761]: "18"
        - generic [ref=e762]: —
        - generic [ref=e763]: long
        - generic [ref=e764]: enough
        - generic [ref=e765]: content
        - generic [ref=e766]: to
        - generic [ref=e767]: force
        - generic [ref=e768]: natural
        - generic [ref=e769]: pagination
        - generic [ref=e770]: across
        - generic [ref=e771]: multiple
        - generic [ref=e772]: pages
        - generic [ref=e773]: so
        - generic [ref=e774]: that
        - generic [ref=e775]: widow
        - generic [ref=e776]: control,
        - generic [ref=e777]: page
        - generic [ref=e778]: fill
        - generic [ref=e779]: and
        - generic [ref=e780]: footers
        - generic [ref=e781]: can
        - generic [ref=e782]: be
        - generic [ref=e783]: verified
        - generic [ref=e784]: visually.
        - generic [ref=e785]: Filler
        - generic [ref=e786]: paragraph
        - generic [ref=e787]: "19"
        - generic [ref=e788]: —
        - generic [ref=e789]: long
        - generic [ref=e790]: enough
        - generic [ref=e791]: content
        - generic [ref=e792]: to
        - generic [ref=e793]: force
        - generic [ref=e794]: natural
        - generic [ref=e795]: pagination
        - generic [ref=e796]: across
        - generic [ref=e797]: multiple
        - generic [ref=e798]: pages
        - generic [ref=e799]: so
        - generic [ref=e800]: that
        - generic [ref=e801]: widow
        - generic [ref=e802]: control,
        - generic [ref=e803]: page
        - generic [ref=e804]: fill
        - generic [ref=e805]: and
        - generic [ref=e806]: footers
        - generic [ref=e807]: can
        - generic [ref=e808]: be
        - generic [ref=e809]: verified
        - generic [ref=e810]: visually.
        - generic [ref=e811]: Filler
        - generic [ref=e812]: paragraph
        - generic [ref=e813]: "20"
        - generic [ref=e814]: —
        - generic [ref=e815]: long
        - generic [ref=e816]: enough
        - generic [ref=e817]: content
        - generic [ref=e818]: to
        - generic [ref=e819]: force
        - generic [ref=e820]: natural
        - generic [ref=e821]: pagination
        - generic [ref=e822]: across
        - generic [ref=e823]: multiple
        - generic [ref=e824]: pages
        - generic [ref=e825]: so
        - generic [ref=e826]: that
        - generic [ref=e827]: widow
        - generic [ref=e828]: control,
        - generic [ref=e829]: page
        - generic [ref=e830]: fill
        - generic [ref=e831]: and
        - generic [ref=e832]: footers
        - generic [ref=e833]: can
        - generic [ref=e834]: be
        - generic [ref=e835]: verified
        - generic [ref=e836]: visually.
        - generic [ref=e837]: Filler
        - generic [ref=e838]: paragraph
        - generic [ref=e839]: "21"
        - generic [ref=e840]: —
        - generic [ref=e841]: long
        - generic [ref=e842]: enough
        - generic [ref=e843]: content
        - generic [ref=e844]: to
        - generic [ref=e845]: force
        - generic [ref=e846]: natural
        - generic [ref=e847]: pagination
        - generic [ref=e848]: across
        - generic [ref=e849]: multiple
        - generic [ref=e850]: pages
        - generic [ref=e851]: so
        - generic [ref=e852]: that
        - generic [ref=e853]: widow
        - generic [ref=e854]: control,
        - generic [ref=e855]: page
        - generic [ref=e856]: fill
        - generic [ref=e857]: and
        - generic [ref=e858]: footers
        - generic [ref=e859]: can
        - generic [ref=e860]: be
        - generic [ref=e861]: verified
        - generic [ref=e862]: visually.
        - generic [ref=e863]: Filler
        - generic [ref=e864]: paragraph
        - generic [ref=e865]: "22"
        - generic [ref=e866]: —
        - generic [ref=e867]: long
        - generic [ref=e868]: enough
        - generic [ref=e869]: content
        - generic [ref=e870]: to
        - generic [ref=e871]: force
        - generic [ref=e872]: natural
        - generic [ref=e873]: pagination
        - generic [ref=e874]: across
        - generic [ref=e875]: multiple
        - generic [ref=e876]: pages
        - generic [ref=e877]: so
        - generic [ref=e878]: that
        - generic [ref=e879]: widow
        - generic [ref=e880]: control,
        - generic [ref=e881]: page
        - generic [ref=e882]: fill
        - generic [ref=e883]: and
        - generic [ref=e884]: footers
        - generic [ref=e885]: can
        - generic [ref=e886]: be
        - generic [ref=e887]: verified
        - generic [ref=e888]: visually.
        - generic [ref=e889]: Filler
        - generic [ref=e890]: paragraph
        - generic [ref=e891]: "23"
        - generic [ref=e892]: —
        - generic [ref=e893]: long
        - generic [ref=e894]: enough
        - generic [ref=e895]: content
        - generic [ref=e896]: to
        - generic [ref=e897]: force
        - generic [ref=e898]: natural
        - generic [ref=e899]: pagination
        - generic [ref=e900]: across
        - generic [ref=e901]: multiple
        - generic [ref=e902]: pages
        - generic [ref=e903]: so
        - generic [ref=e904]: that
        - generic [ref=e905]: widow
        - generic [ref=e906]: control,
        - generic [ref=e907]: page
        - generic [ref=e908]: fill
        - generic [ref=e909]: and
        - generic [ref=e910]: footers
        - generic [ref=e911]: can
        - generic [ref=e912]: be
        - generic [ref=e913]: verified
        - generic [ref=e914]: visually.
        - generic [ref=e915]: Filler
        - generic [ref=e916]: paragraph
        - generic [ref=e917]: "24"
        - generic [ref=e918]: —
        - generic [ref=e919]: long
        - generic [ref=e920]: enough
        - generic [ref=e921]: content
        - generic [ref=e922]: to
        - generic [ref=e923]: force
        - generic [ref=e924]: natural
        - generic [ref=e925]: pagination
        - generic [ref=e926]: across
        - generic [ref=e927]: multiple
        - generic [ref=e928]: pages
        - generic [ref=e929]: so
        - generic [ref=e930]: that
        - generic [ref=e931]: widow
        - generic [ref=e932]: control,
        - generic [ref=e933]: page
        - generic [ref=e934]: fill
        - generic [ref=e935]: and
        - generic [ref=e936]: footers
        - generic [ref=e937]: can
        - generic [ref=e938]: be
        - generic [ref=e939]: verified
        - generic [ref=e940]: visually.
        - generic [ref=e941]: DocxInWeb
        - generic [ref=e942]: Fidelity
        - generic [ref=e943]: Sample
        - generic [ref=e944]: Page
        - generic [ref=e945]: "2"
        - generic [ref=e946]: of
        - generic [ref=e947]: "4"
      - generic [ref=e949]:
        - generic [ref=e950]: Filler
        - generic [ref=e951]: paragraph
        - generic [ref=e952]: "25"
        - generic [ref=e953]: —
        - generic [ref=e954]: long
        - generic [ref=e955]: enough
        - generic [ref=e956]: content
        - generic [ref=e957]: to
        - generic [ref=e958]: force
        - generic [ref=e959]: natural
        - generic [ref=e960]: pagination
        - generic [ref=e961]: across
        - generic [ref=e962]: multiple
        - generic [ref=e963]: pages
        - generic [ref=e964]: so
        - generic [ref=e965]: that
        - generic [ref=e966]: widow
        - generic [ref=e967]: control,
        - generic [ref=e968]: page
        - generic [ref=e969]: fill
        - generic [ref=e970]: and
        - generic [ref=e971]: footers
        - generic [ref=e972]: can
        - generic [ref=e973]: be
        - generic [ref=e974]: verified
        - generic [ref=e975]: visually.
        - generic [ref=e976]: Filler
        - generic [ref=e977]: paragraph
        - generic [ref=e978]: "26"
        - generic [ref=e979]: —
        - generic [ref=e980]: long
        - generic [ref=e981]: enough
        - generic [ref=e982]: content
        - generic [ref=e983]: to
        - generic [ref=e984]: force
        - generic [ref=e985]: natural
        - generic [ref=e986]: pagination
        - generic [ref=e987]: across
        - generic [ref=e988]: multiple
        - generic [ref=e989]: pages
        - generic [ref=e990]: so
        - generic [ref=e991]: that
        - generic [ref=e992]: widow
        - generic [ref=e993]: control,
        - generic [ref=e994]: page
        - generic [ref=e995]: fill
        - generic [ref=e996]: and
        - generic [ref=e997]: footers
        - generic [ref=e998]: can
        - generic [ref=e999]: be
        - generic [ref=e1000]: verified
        - generic [ref=e1001]: visually.
        - generic [ref=e1002]: Filler
        - generic [ref=e1003]: paragraph
        - generic [ref=e1004]: "27"
        - generic [ref=e1005]: —
        - generic [ref=e1006]: long
        - generic [ref=e1007]: enough
        - generic [ref=e1008]: content
        - generic [ref=e1009]: to
        - generic [ref=e1010]: force
        - generic [ref=e1011]: natural
        - generic [ref=e1012]: pagination
        - generic [ref=e1013]: across
        - generic [ref=e1014]: multiple
        - generic [ref=e1015]: pages
        - generic [ref=e1016]: so
        - generic [ref=e1017]: that
        - generic [ref=e1018]: widow
        - generic [ref=e1019]: control,
        - generic [ref=e1020]: page
        - generic [ref=e1021]: fill
        - generic [ref=e1022]: and
        - generic [ref=e1023]: footers
        - generic [ref=e1024]: can
        - generic [ref=e1025]: be
        - generic [ref=e1026]: verified
        - generic [ref=e1027]: visually.
        - generic [ref=e1028]: Filler
        - generic [ref=e1029]: paragraph
        - generic [ref=e1030]: "28"
        - generic [ref=e1031]: —
        - generic [ref=e1032]: long
        - generic [ref=e1033]: enough
        - generic [ref=e1034]: content
        - generic [ref=e1035]: to
        - generic [ref=e1036]: force
        - generic [ref=e1037]: natural
        - generic [ref=e1038]: pagination
        - generic [ref=e1039]: across
        - generic [ref=e1040]: multiple
        - generic [ref=e1041]: pages
        - generic [ref=e1042]: so
        - generic [ref=e1043]: that
        - generic [ref=e1044]: widow
        - generic [ref=e1045]: control,
        - generic [ref=e1046]: page
        - generic [ref=e1047]: fill
        - generic [ref=e1048]: and
        - generic [ref=e1049]: footers
        - generic [ref=e1050]: can
        - generic [ref=e1051]: be
        - generic [ref=e1052]: verified
        - generic [ref=e1053]: visually.
        - generic [ref=e1054]: Filler
        - generic [ref=e1055]: paragraph
        - generic [ref=e1056]: "29"
        - generic [ref=e1057]: —
        - generic [ref=e1058]: long
        - generic [ref=e1059]: enough
        - generic [ref=e1060]: content
        - generic [ref=e1061]: to
        - generic [ref=e1062]: force
        - generic [ref=e1063]: natural
        - generic [ref=e1064]: pagination
        - generic [ref=e1065]: across
        - generic [ref=e1066]: multiple
        - generic [ref=e1067]: pages
        - generic [ref=e1068]: so
        - generic [ref=e1069]: that
        - generic [ref=e1070]: widow
        - generic [ref=e1071]: control,
        - generic [ref=e1072]: page
        - generic [ref=e1073]: fill
        - generic [ref=e1074]: and
        - generic [ref=e1075]: footers
        - generic [ref=e1076]: can
        - generic [ref=e1077]: be
        - generic [ref=e1078]: verified
        - generic [ref=e1079]: visually.
        - generic [ref=e1080]: Filler
        - generic [ref=e1081]: paragraph
        - generic [ref=e1082]: "30"
        - generic [ref=e1083]: —
        - generic [ref=e1084]: long
        - generic [ref=e1085]: enough
        - generic [ref=e1086]: content
        - generic [ref=e1087]: to
        - generic [ref=e1088]: force
        - generic [ref=e1089]: natural
        - generic [ref=e1090]: pagination
        - generic [ref=e1091]: across
        - generic [ref=e1092]: multiple
        - generic [ref=e1093]: pages
        - generic [ref=e1094]: so
        - generic [ref=e1095]: that
        - generic [ref=e1096]: widow
        - generic [ref=e1097]: control,
        - generic [ref=e1098]: page
        - generic [ref=e1099]: fill
        - generic [ref=e1100]: and
        - generic [ref=e1101]: footers
        - generic [ref=e1102]: can
        - generic [ref=e1103]: be
        - generic [ref=e1104]: verified
        - generic [ref=e1105]: visually.
        - generic [ref=e1106]: Filler
        - generic [ref=e1107]: paragraph
        - generic [ref=e1108]: "31"
        - generic [ref=e1109]: —
        - generic [ref=e1110]: long
        - generic [ref=e1111]: enough
        - generic [ref=e1112]: content
        - generic [ref=e1113]: to
        - generic [ref=e1114]: force
        - generic [ref=e1115]: natural
        - generic [ref=e1116]: pagination
        - generic [ref=e1117]: across
        - generic [ref=e1118]: multiple
        - generic [ref=e1119]: pages
        - generic [ref=e1120]: so
        - generic [ref=e1121]: that
        - generic [ref=e1122]: widow
        - generic [ref=e1123]: control,
        - generic [ref=e1124]: page
        - generic [ref=e1125]: fill
        - generic [ref=e1126]: and
        - generic [ref=e1127]: footers
        - generic [ref=e1128]: can
        - generic [ref=e1129]: be
        - generic [ref=e1130]: verified
        - generic [ref=e1131]: visually.
        - generic [ref=e1132]: Filler
        - generic [ref=e1133]: paragraph
        - generic [ref=e1134]: "32"
        - generic [ref=e1135]: —
        - generic [ref=e1136]: long
        - generic [ref=e1137]: enough
        - generic [ref=e1138]: content
        - generic [ref=e1139]: to
        - generic [ref=e1140]: force
        - generic [ref=e1141]: natural
        - generic [ref=e1142]: pagination
        - generic [ref=e1143]: across
        - generic [ref=e1144]: multiple
        - generic [ref=e1145]: pages
        - generic [ref=e1146]: so
        - generic [ref=e1147]: that
        - generic [ref=e1148]: widow
        - generic [ref=e1149]: control,
        - generic [ref=e1150]: page
        - generic [ref=e1151]: fill
        - generic [ref=e1152]: and
        - generic [ref=e1153]: footers
        - generic [ref=e1154]: can
        - generic [ref=e1155]: be
        - generic [ref=e1156]: verified
        - generic [ref=e1157]: visually.
        - generic [ref=e1158]: Filler
        - generic [ref=e1159]: paragraph
        - generic [ref=e1160]: "33"
        - generic [ref=e1161]: —
        - generic [ref=e1162]: long
        - generic [ref=e1163]: enough
        - generic [ref=e1164]: content
        - generic [ref=e1165]: to
        - generic [ref=e1166]: force
        - generic [ref=e1167]: natural
        - generic [ref=e1168]: pagination
        - generic [ref=e1169]: across
        - generic [ref=e1170]: multiple
        - generic [ref=e1171]: pages
        - generic [ref=e1172]: so
        - generic [ref=e1173]: that
        - generic [ref=e1174]: widow
        - generic [ref=e1175]: control,
        - generic [ref=e1176]: page
        - generic [ref=e1177]: fill
        - generic [ref=e1178]: and
        - generic [ref=e1179]: footers
        - generic [ref=e1180]: can
        - generic [ref=e1181]: be
        - generic [ref=e1182]: verified
        - generic [ref=e1183]: visually.
        - generic [ref=e1184]: Filler
        - generic [ref=e1185]: paragraph
        - generic [ref=e1186]: "34"
        - generic [ref=e1187]: —
        - generic [ref=e1188]: long
        - generic [ref=e1189]: enough
        - generic [ref=e1190]: content
        - generic [ref=e1191]: to
        - generic [ref=e1192]: force
        - generic [ref=e1193]: natural
        - generic [ref=e1194]: pagination
        - generic [ref=e1195]: across
        - generic [ref=e1196]: multiple
        - generic [ref=e1197]: pages
        - generic [ref=e1198]: so
        - generic [ref=e1199]: that
        - generic [ref=e1200]: widow
        - generic [ref=e1201]: control,
        - generic [ref=e1202]: page
        - generic [ref=e1203]: fill
        - generic [ref=e1204]: and
        - generic [ref=e1205]: footers
        - generic [ref=e1206]: can
        - generic [ref=e1207]: be
        - generic [ref=e1208]: verified
        - generic [ref=e1209]: visually.
        - generic [ref=e1210]: Filler
        - generic [ref=e1211]: paragraph
        - generic [ref=e1212]: "35"
        - generic [ref=e1213]: —
        - generic [ref=e1214]: long
        - generic [ref=e1215]: enough
        - generic [ref=e1216]: content
        - generic [ref=e1217]: to
        - generic [ref=e1218]: force
        - generic [ref=e1219]: natural
        - generic [ref=e1220]: pagination
        - generic [ref=e1221]: across
        - generic [ref=e1222]: multiple
        - generic [ref=e1223]: pages
        - generic [ref=e1224]: so
        - generic [ref=e1225]: that
        - generic [ref=e1226]: widow
        - generic [ref=e1227]: control,
        - generic [ref=e1228]: page
        - generic [ref=e1229]: fill
        - generic [ref=e1230]: and
        - generic [ref=e1231]: footers
        - generic [ref=e1232]: can
        - generic [ref=e1233]: be
        - generic [ref=e1234]: verified
        - generic [ref=e1235]: visually.
        - generic [ref=e1236]: Filler
        - generic [ref=e1237]: paragraph
        - generic [ref=e1238]: "36"
        - generic [ref=e1239]: —
        - generic [ref=e1240]: long
        - generic [ref=e1241]: enough
        - generic [ref=e1242]: content
        - generic [ref=e1243]: to
        - generic [ref=e1244]: force
        - generic [ref=e1245]: natural
        - generic [ref=e1246]: pagination
        - generic [ref=e1247]: across
        - generic [ref=e1248]: multiple
        - generic [ref=e1249]: pages
        - generic [ref=e1250]: so
        - generic [ref=e1251]: that
        - generic [ref=e1252]: widow
        - generic [ref=e1253]: control,
        - generic [ref=e1254]: page
        - generic [ref=e1255]: fill
        - generic [ref=e1256]: and
        - generic [ref=e1257]: footers
        - generic [ref=e1258]: can
        - generic [ref=e1259]: be
        - generic [ref=e1260]: verified
        - generic [ref=e1261]: visually.
        - generic [ref=e1262]: Filler
        - generic [ref=e1263]: paragraph
        - generic [ref=e1264]: "37"
        - generic [ref=e1265]: —
        - generic [ref=e1266]: long
        - generic [ref=e1267]: enough
        - generic [ref=e1268]: content
        - generic [ref=e1269]: to
        - generic [ref=e1270]: force
        - generic [ref=e1271]: natural
        - generic [ref=e1272]: pagination
        - generic [ref=e1273]: across
        - generic [ref=e1274]: multiple
        - generic [ref=e1275]: pages
        - generic [ref=e1276]: so
        - generic [ref=e1277]: that
        - generic [ref=e1278]: widow
        - generic [ref=e1279]: control,
        - generic [ref=e1280]: page
        - generic [ref=e1281]: fill
        - generic [ref=e1282]: and
        - generic [ref=e1283]: footers
        - generic [ref=e1284]: can
        - generic [ref=e1285]: be
        - generic [ref=e1286]: verified
        - generic [ref=e1287]: visually.
        - generic [ref=e1288]: Filler
        - generic [ref=e1289]: paragraph
        - generic [ref=e1290]: "38"
        - generic [ref=e1291]: —
        - generic [ref=e1292]: long
        - generic [ref=e1293]: enough
        - generic [ref=e1294]: content
        - generic [ref=e1295]: to
        - generic [ref=e1296]: force
        - generic [ref=e1297]: natural
        - generic [ref=e1298]: pagination
        - generic [ref=e1299]: across
        - generic [ref=e1300]: multiple
        - generic [ref=e1301]: pages
        - generic [ref=e1302]: so
        - generic [ref=e1303]: that
        - generic [ref=e1304]: widow
        - generic [ref=e1305]: control,
        - generic [ref=e1306]: page
        - generic [ref=e1307]: fill
        - generic [ref=e1308]: and
        - generic [ref=e1309]: footers
        - generic [ref=e1310]: can
        - generic [ref=e1311]: be
        - generic [ref=e1312]: verified
        - generic [ref=e1313]: visually.
        - generic [ref=e1314]: Filler
        - generic [ref=e1315]: paragraph
        - generic [ref=e1316]: "39"
        - generic [ref=e1317]: —
        - generic [ref=e1318]: long
        - generic [ref=e1319]: enough
        - generic [ref=e1320]: content
        - generic [ref=e1321]: to
        - generic [ref=e1322]: force
        - generic [ref=e1323]: natural
        - generic [ref=e1324]: pagination
        - generic [ref=e1325]: across
        - generic [ref=e1326]: multiple
        - generic [ref=e1327]: pages
        - generic [ref=e1328]: so
        - generic [ref=e1329]: that
        - generic [ref=e1330]: widow
        - generic [ref=e1331]: control,
        - generic [ref=e1332]: page
        - generic [ref=e1333]: fill
        - generic [ref=e1334]: and
        - generic [ref=e1335]: footers
        - generic [ref=e1336]: can
        - generic [ref=e1337]: be
        - generic [ref=e1338]: verified
        - generic [ref=e1339]: visually.
        - generic [ref=e1340]: Filler
        - generic [ref=e1341]: paragraph
        - generic [ref=e1342]: "40"
        - generic [ref=e1343]: —
        - generic [ref=e1344]: long
        - generic [ref=e1345]: enough
        - generic [ref=e1346]: content
        - generic [ref=e1347]: to
        - generic [ref=e1348]: force
        - generic [ref=e1349]: natural
        - generic [ref=e1350]: pagination
        - generic [ref=e1351]: across
        - generic [ref=e1352]: multiple
        - generic [ref=e1353]: pages
        - generic [ref=e1354]: so
        - generic [ref=e1355]: that
        - generic [ref=e1356]: widow
        - generic [ref=e1357]: control,
        - generic [ref=e1358]: page
        - generic [ref=e1359]: fill
        - generic [ref=e1360]: and
        - generic [ref=e1361]: footers
        - generic [ref=e1362]: can
        - generic [ref=e1363]: be
        - generic [ref=e1364]: verified
        - generic [ref=e1365]: visually.
        - generic [ref=e1366]: Filler
        - generic [ref=e1367]: paragraph
        - generic [ref=e1368]: "41"
        - generic [ref=e1369]: —
        - generic [ref=e1370]: long
        - generic [ref=e1371]: enough
        - generic [ref=e1372]: content
        - generic [ref=e1373]: to
        - generic [ref=e1374]: force
        - generic [ref=e1375]: natural
        - generic [ref=e1376]: pagination
        - generic [ref=e1377]: across
        - generic [ref=e1378]: multiple
        - generic [ref=e1379]: pages
        - generic [ref=e1380]: so
        - generic [ref=e1381]: that
        - generic [ref=e1382]: widow
        - generic [ref=e1383]: control,
        - generic [ref=e1384]: page
        - generic [ref=e1385]: fill
        - generic [ref=e1386]: and
        - generic [ref=e1387]: footers
        - generic [ref=e1388]: can
        - generic [ref=e1389]: be
        - generic [ref=e1390]: verified
        - generic [ref=e1391]: visually.
        - generic [ref=e1392]: Filler
        - generic [ref=e1393]: paragraph
        - generic [ref=e1394]: "42"
        - generic [ref=e1395]: —
        - generic [ref=e1396]: long
        - generic [ref=e1397]: enough
        - generic [ref=e1398]: content
        - generic [ref=e1399]: to
        - generic [ref=e1400]: force
        - generic [ref=e1401]: natural
        - generic [ref=e1402]: pagination
        - generic [ref=e1403]: across
        - generic [ref=e1404]: multiple
        - generic [ref=e1405]: pages
        - generic [ref=e1406]: so
        - generic [ref=e1407]: that
        - generic [ref=e1408]: widow
        - generic [ref=e1409]: control,
        - generic [ref=e1410]: page
        - generic [ref=e1411]: fill
        - generic [ref=e1412]: and
        - generic [ref=e1413]: footers
        - generic [ref=e1414]: can
        - generic [ref=e1415]: be
        - generic [ref=e1416]: verified
        - generic [ref=e1417]: visually.
        - generic [ref=e1418]: Filler
        - generic [ref=e1419]: paragraph
        - generic [ref=e1420]: "43"
        - generic [ref=e1421]: —
        - generic [ref=e1422]: long
        - generic [ref=e1423]: enough
        - generic [ref=e1424]: content
        - generic [ref=e1425]: to
        - generic [ref=e1426]: force
        - generic [ref=e1427]: natural
        - generic [ref=e1428]: pagination
        - generic [ref=e1429]: across
        - generic [ref=e1430]: multiple
        - generic [ref=e1431]: pages
        - generic [ref=e1432]: so
        - generic [ref=e1433]: that
        - generic [ref=e1434]: widow
        - generic [ref=e1435]: control,
        - generic [ref=e1436]: page
        - generic [ref=e1437]: fill
        - generic [ref=e1438]: and
        - generic [ref=e1439]: footers
        - generic [ref=e1440]: can
        - generic [ref=e1441]: be
        - generic [ref=e1442]: verified
        - generic [ref=e1443]: visually.
        - generic [ref=e1444]: Filler
        - generic [ref=e1445]: paragraph
        - generic [ref=e1446]: "44"
        - generic [ref=e1447]: —
        - generic [ref=e1448]: long
        - generic [ref=e1449]: enough
        - generic [ref=e1450]: content
        - generic [ref=e1451]: to
        - generic [ref=e1452]: force
        - generic [ref=e1453]: natural
        - generic [ref=e1454]: pagination
        - generic [ref=e1455]: across
        - generic [ref=e1456]: multiple
        - generic [ref=e1457]: pages
        - generic [ref=e1458]: so
        - generic [ref=e1459]: that
        - generic [ref=e1460]: widow
        - generic [ref=e1461]: control,
        - generic [ref=e1462]: page
        - generic [ref=e1463]: fill
        - generic [ref=e1464]: and
        - generic [ref=e1465]: footers
        - generic [ref=e1466]: can
        - generic [ref=e1467]: be
        - generic [ref=e1468]: verified
        - generic [ref=e1469]: visually.
        - generic [ref=e1470]: Filler
        - generic [ref=e1471]: paragraph
        - generic [ref=e1472]: "45"
        - generic [ref=e1473]: —
        - generic [ref=e1474]: long
        - generic [ref=e1475]: enough
        - generic [ref=e1476]: content
        - generic [ref=e1477]: to
        - generic [ref=e1478]: force
        - generic [ref=e1479]: natural
        - generic [ref=e1480]: pagination
        - generic [ref=e1481]: across
        - generic [ref=e1482]: multiple
        - generic [ref=e1483]: pages
        - generic [ref=e1484]: so
        - generic [ref=e1485]: that
        - generic [ref=e1486]: widow
        - generic [ref=e1487]: control,
        - generic [ref=e1488]: page
        - generic [ref=e1489]: fill
        - generic [ref=e1490]: and
        - generic [ref=e1491]: footers
        - generic [ref=e1492]: can
        - generic [ref=e1493]: be
        - generic [ref=e1494]: verified
        - generic [ref=e1495]: visually.
        - generic [ref=e1496]: Filler
        - generic [ref=e1497]: paragraph
        - generic [ref=e1498]: "46"
        - generic [ref=e1499]: —
        - generic [ref=e1500]: long
        - generic [ref=e1501]: enough
        - generic [ref=e1502]: content
        - generic [ref=e1503]: to
        - generic [ref=e1504]: force
        - generic [ref=e1505]: natural
        - generic [ref=e1506]: pagination
        - generic [ref=e1507]: across
        - generic [ref=e1508]: multiple
        - generic [ref=e1509]: pages
        - generic [ref=e1510]: so
        - generic [ref=e1511]: that
        - generic [ref=e1512]: widow
        - generic [ref=e1513]: control,
        - generic [ref=e1514]: page
        - generic [ref=e1515]: fill
        - generic [ref=e1516]: and
        - generic [ref=e1517]: footers
        - generic [ref=e1518]: can
        - generic [ref=e1519]: be
        - generic [ref=e1520]: verified
        - generic [ref=e1521]: visually.
        - generic [ref=e1522]: Filler
        - generic [ref=e1523]: paragraph
        - generic [ref=e1524]: "47"
        - generic [ref=e1525]: —
        - generic [ref=e1526]: long
        - generic [ref=e1527]: enough
        - generic [ref=e1528]: content
        - generic [ref=e1529]: to
        - generic [ref=e1530]: force
        - generic [ref=e1531]: natural
        - generic [ref=e1532]: pagination
        - generic [ref=e1533]: across
        - generic [ref=e1534]: multiple
        - generic [ref=e1535]: pages
        - generic [ref=e1536]: so
        - generic [ref=e1537]: that
        - generic [ref=e1538]: widow
        - generic [ref=e1539]: control,
        - generic [ref=e1540]: page
        - generic [ref=e1541]: fill
        - generic [ref=e1542]: and
        - generic [ref=e1543]: footers
        - generic [ref=e1544]: can
        - generic [ref=e1545]: be
        - generic [ref=e1546]: verified
        - generic [ref=e1547]: visually.
        - generic [ref=e1548]: Filler
        - generic [ref=e1549]: paragraph
        - generic [ref=e1550]: "48"
        - generic [ref=e1551]: —
        - generic [ref=e1552]: long
        - generic [ref=e1553]: enough
        - generic [ref=e1554]: content
        - generic [ref=e1555]: to
        - generic [ref=e1556]: force
        - generic [ref=e1557]: natural
        - generic [ref=e1558]: pagination
        - generic [ref=e1559]: across
        - generic [ref=e1560]: multiple
        - generic [ref=e1561]: pages
        - generic [ref=e1562]: so
        - generic [ref=e1563]: that
        - generic [ref=e1564]: widow
        - generic [ref=e1565]: control,
        - generic [ref=e1566]: page
        - generic [ref=e1567]: fill
        - generic [ref=e1568]: and
        - generic [ref=e1569]: footers
        - generic [ref=e1570]: can
        - generic [ref=e1571]: be
        - generic [ref=e1572]: verified
        - generic [ref=e1573]: visually.
        - generic [ref=e1574]: Filler
        - generic [ref=e1575]: paragraph
        - generic [ref=e1576]: "49"
        - generic [ref=e1577]: —
        - generic [ref=e1578]: long
        - generic [ref=e1579]: enough
        - generic [ref=e1580]: content
        - generic [ref=e1581]: to
        - generic [ref=e1582]: force
        - generic [ref=e1583]: natural
        - generic [ref=e1584]: pagination
        - generic [ref=e1585]: across
        - generic [ref=e1586]: multiple
        - generic [ref=e1587]: pages
        - generic [ref=e1588]: so
        - generic [ref=e1589]: that
        - generic [ref=e1590]: widow
        - generic [ref=e1591]: control,
        - generic [ref=e1592]: page
        - generic [ref=e1593]: fill
        - generic [ref=e1594]: and
        - generic [ref=e1595]: footers
        - generic [ref=e1596]: can
        - generic [ref=e1597]: be
        - generic [ref=e1598]: verified
        - generic [ref=e1599]: visually.
        - generic [ref=e1600]: DocxInWeb
        - generic [ref=e1601]: Fidelity
        - generic [ref=e1602]: Sample
        - generic [ref=e1603]: Page
        - generic [ref=e1604]: "3"
        - generic [ref=e1605]: of
        - generic [ref=e1606]: "4"
      - generic [ref=e1608]:
        - generic [ref=e1609]: Filler
        - generic [ref=e1610]: paragraph
        - generic [ref=e1611]: "50"
        - generic [ref=e1612]: —
        - generic [ref=e1613]: long
        - generic [ref=e1614]: enough
        - generic [ref=e1615]: content
        - generic [ref=e1616]: to
        - generic [ref=e1617]: force
        - generic [ref=e1618]: natural
        - generic [ref=e1619]: pagination
        - generic [ref=e1620]: across
        - generic [ref=e1621]: multiple
        - generic [ref=e1622]: pages
        - generic [ref=e1623]: so
        - generic [ref=e1624]: that
        - generic [ref=e1625]: widow
        - generic [ref=e1626]: control,
        - generic [ref=e1627]: page
        - generic [ref=e1628]: fill
        - generic [ref=e1629]: and
        - generic [ref=e1630]: footers
        - generic [ref=e1631]: can
        - generic [ref=e1632]: be
        - generic [ref=e1633]: verified
        - generic [ref=e1634]: visually.
        - generic [ref=e1635]: Filler
        - generic [ref=e1636]: paragraph
        - generic [ref=e1637]: "51"
        - generic [ref=e1638]: —
        - generic [ref=e1639]: long
        - generic [ref=e1640]: enough
        - generic [ref=e1641]: content
        - generic [ref=e1642]: to
        - generic [ref=e1643]: force
        - generic [ref=e1644]: natural
        - generic [ref=e1645]: pagination
        - generic [ref=e1646]: across
        - generic [ref=e1647]: multiple
        - generic [ref=e1648]: pages
        - generic [ref=e1649]: so
        - generic [ref=e1650]: that
        - generic [ref=e1651]: widow
        - generic [ref=e1652]: control,
        - generic [ref=e1653]: page
        - generic [ref=e1654]: fill
        - generic [ref=e1655]: and
        - generic [ref=e1656]: footers
        - generic [ref=e1657]: can
        - generic [ref=e1658]: be
        - generic [ref=e1659]: verified
        - generic [ref=e1660]: visually.
        - generic [ref=e1661]: Filler
        - generic [ref=e1662]: paragraph
        - generic [ref=e1663]: "52"
        - generic [ref=e1664]: —
        - generic [ref=e1665]: long
        - generic [ref=e1666]: enough
        - generic [ref=e1667]: content
        - generic [ref=e1668]: to
        - generic [ref=e1669]: force
        - generic [ref=e1670]: natural
        - generic [ref=e1671]: pagination
        - generic [ref=e1672]: across
        - generic [ref=e1673]: multiple
        - generic [ref=e1674]: pages
        - generic [ref=e1675]: so
        - generic [ref=e1676]: that
        - generic [ref=e1677]: widow
        - generic [ref=e1678]: control,
        - generic [ref=e1679]: page
        - generic [ref=e1680]: fill
        - generic [ref=e1681]: and
        - generic [ref=e1682]: footers
        - generic [ref=e1683]: can
        - generic [ref=e1684]: be
        - generic [ref=e1685]: verified
        - generic [ref=e1686]: visually.
        - generic [ref=e1687]: Filler
        - generic [ref=e1688]: paragraph
        - generic [ref=e1689]: "53"
        - generic [ref=e1690]: —
        - generic [ref=e1691]: long
        - generic [ref=e1692]: enough
        - generic [ref=e1693]: content
        - generic [ref=e1694]: to
        - generic [ref=e1695]: force
        - generic [ref=e1696]: natural
        - generic [ref=e1697]: pagination
        - generic [ref=e1698]: across
        - generic [ref=e1699]: multiple
        - generic [ref=e1700]: pages
        - generic [ref=e1701]: so
        - generic [ref=e1702]: that
        - generic [ref=e1703]: widow
        - generic [ref=e1704]: control,
        - generic [ref=e1705]: page
        - generic [ref=e1706]: fill
        - generic [ref=e1707]: and
        - generic [ref=e1708]: footers
        - generic [ref=e1709]: can
        - generic [ref=e1710]: be
        - generic [ref=e1711]: verified
        - generic [ref=e1712]: visually.
        - generic [ref=e1713]: Filler
        - generic [ref=e1714]: paragraph
        - generic [ref=e1715]: "54"
        - generic [ref=e1716]: —
        - generic [ref=e1717]: long
        - generic [ref=e1718]: enough
        - generic [ref=e1719]: content
        - generic [ref=e1720]: to
        - generic [ref=e1721]: force
        - generic [ref=e1722]: natural
        - generic [ref=e1723]: pagination
        - generic [ref=e1724]: across
        - generic [ref=e1725]: multiple
        - generic [ref=e1726]: pages
        - generic [ref=e1727]: so
        - generic [ref=e1728]: that
        - generic [ref=e1729]: widow
        - generic [ref=e1730]: control,
        - generic [ref=e1731]: page
        - generic [ref=e1732]: fill
        - generic [ref=e1733]: and
        - generic [ref=e1734]: footers
        - generic [ref=e1735]: can
        - generic [ref=e1736]: be
        - generic [ref=e1737]: verified
        - generic [ref=e1738]: visually.
        - generic [ref=e1739]: Filler
        - generic [ref=e1740]: paragraph
        - generic [ref=e1741]: "55"
        - generic [ref=e1742]: —
        - generic [ref=e1743]: long
        - generic [ref=e1744]: enough
        - generic [ref=e1745]: content
        - generic [ref=e1746]: to
        - generic [ref=e1747]: force
        - generic [ref=e1748]: natural
        - generic [ref=e1749]: pagination
        - generic [ref=e1750]: across
        - generic [ref=e1751]: multiple
        - generic [ref=e1752]: pages
        - generic [ref=e1753]: so
        - generic [ref=e1754]: that
        - generic [ref=e1755]: widow
        - generic [ref=e1756]: control,
        - generic [ref=e1757]: page
        - generic [ref=e1758]: fill
        - generic [ref=e1759]: and
        - generic [ref=e1760]: footers
        - generic [ref=e1761]: can
        - generic [ref=e1762]: be
        - generic [ref=e1763]: verified
        - generic [ref=e1764]: visually.
        - generic [ref=e1765]: Filler
        - generic [ref=e1766]: paragraph
        - generic [ref=e1767]: "56"
        - generic [ref=e1768]: —
        - generic [ref=e1769]: long
        - generic [ref=e1770]: enough
        - generic [ref=e1771]: content
        - generic [ref=e1772]: to
        - generic [ref=e1773]: force
        - generic [ref=e1774]: natural
        - generic [ref=e1775]: pagination
        - generic [ref=e1776]: across
        - generic [ref=e1777]: multiple
        - generic [ref=e1778]: pages
        - generic [ref=e1779]: so
        - generic [ref=e1780]: that
        - generic [ref=e1781]: widow
        - generic [ref=e1782]: control,
        - generic [ref=e1783]: page
        - generic [ref=e1784]: fill
        - generic [ref=e1785]: and
        - generic [ref=e1786]: footers
        - generic [ref=e1787]: can
        - generic [ref=e1788]: be
        - generic [ref=e1789]: verified
        - generic [ref=e1790]: visually.
        - generic [ref=e1791]: Filler
        - generic [ref=e1792]: paragraph
        - generic [ref=e1793]: "57"
        - generic [ref=e1794]: —
        - generic [ref=e1795]: long
        - generic [ref=e1796]: enough
        - generic [ref=e1797]: content
        - generic [ref=e1798]: to
        - generic [ref=e1799]: force
        - generic [ref=e1800]: natural
        - generic [ref=e1801]: pagination
        - generic [ref=e1802]: across
        - generic [ref=e1803]: multiple
        - generic [ref=e1804]: pages
        - generic [ref=e1805]: so
        - generic [ref=e1806]: that
        - generic [ref=e1807]: widow
        - generic [ref=e1808]: control,
        - generic [ref=e1809]: page
        - generic [ref=e1810]: fill
        - generic [ref=e1811]: and
        - generic [ref=e1812]: footers
        - generic [ref=e1813]: can
        - generic [ref=e1814]: be
        - generic [ref=e1815]: verified
        - generic [ref=e1816]: visually.
        - generic [ref=e1817]: Filler
        - generic [ref=e1818]: paragraph
        - generic [ref=e1819]: "58"
        - generic [ref=e1820]: —
        - generic [ref=e1821]: long
        - generic [ref=e1822]: enough
        - generic [ref=e1823]: content
        - generic [ref=e1824]: to
        - generic [ref=e1825]: force
        - generic [ref=e1826]: natural
        - generic [ref=e1827]: pagination
        - generic [ref=e1828]: across
        - generic [ref=e1829]: multiple
        - generic [ref=e1830]: pages
        - generic [ref=e1831]: so
        - generic [ref=e1832]: that
        - generic [ref=e1833]: widow
        - generic [ref=e1834]: control,
        - generic [ref=e1835]: page
        - generic [ref=e1836]: fill
        - generic [ref=e1837]: and
        - generic [ref=e1838]: footers
        - generic [ref=e1839]: can
        - generic [ref=e1840]: be
        - generic [ref=e1841]: verified
        - generic [ref=e1842]: visually.
        - generic [ref=e1843]: Filler
        - generic [ref=e1844]: paragraph
        - generic [ref=e1845]: "59"
        - generic [ref=e1846]: —
        - generic [ref=e1847]: long
        - generic [ref=e1848]: enough
        - generic [ref=e1849]: content
        - generic [ref=e1850]: to
        - generic [ref=e1851]: force
        - generic [ref=e1852]: natural
        - generic [ref=e1853]: pagination
        - generic [ref=e1854]: across
        - generic [ref=e1855]: multiple
        - generic [ref=e1856]: pages
        - generic [ref=e1857]: so
        - generic [ref=e1858]: that
        - generic [ref=e1859]: widow
        - generic [ref=e1860]: control,
        - generic [ref=e1861]: page
        - generic [ref=e1862]: fill
        - generic [ref=e1863]: and
        - generic [ref=e1864]: footers
        - generic [ref=e1865]: can
        - generic [ref=e1866]: be
        - generic [ref=e1867]: verified
        - generic [ref=e1868]: visually.
        - generic [ref=e1869]: Filler
        - generic [ref=e1870]: paragraph
        - generic [ref=e1871]: "60"
        - generic [ref=e1872]: —
        - generic [ref=e1873]: long
        - generic [ref=e1874]: enough
        - generic [ref=e1875]: content
        - generic [ref=e1876]: to
        - generic [ref=e1877]: force
        - generic [ref=e1878]: natural
        - generic [ref=e1879]: pagination
        - generic [ref=e1880]: across
        - generic [ref=e1881]: multiple
        - generic [ref=e1882]: pages
        - generic [ref=e1883]: so
        - generic [ref=e1884]: that
        - generic [ref=e1885]: widow
        - generic [ref=e1886]: control,
        - generic [ref=e1887]: page
        - generic [ref=e1888]: fill
        - generic [ref=e1889]: and
        - generic [ref=e1890]: footers
        - generic [ref=e1891]: can
        - generic [ref=e1892]: be
        - generic [ref=e1893]: verified
        - generic [ref=e1894]: visually.
        - generic [ref=e1895]: DocxInWeb
        - generic [ref=e1896]: Fidelity
        - generic [ref=e1897]: Sample
        - generic [ref=e1898]: Page
        - generic [ref=e1899]: "4"
        - generic [ref=e1900]: of
        - generic [ref=e1901]: "4"
      - generic [ref=e1902]:
        - generic [ref=e1903]:
          - generic [ref=e1904]: AR
          - generic [ref=e1905]:
            - generic [ref=e1906]: Ada Reviewer
            - generic [ref=e1907]: 6/1/2026
        - generic [ref=e1908]: Should this be brand blue instead of red?
        - textbox "Reply…" [ref=e1909]
      - generic [ref=e1910]:
        - generic [ref=e1911]:
          - generic [ref=e1912]: BE
          - generic [ref=e1913]:
            - generic [ref=e1914]: Bob Editor
            - generic [ref=e1915]: 6/2/2026
        - generic [ref=e1916]: Numbering restarts here — double-check the list level.
        - textbox "Reply…" [ref=e1917]
```

# Test source

```ts
  132 |     await page.mouse.dblclick(box.x + box.width / 2, box.y + box.height / 2);
  133 |     await page.waitForTimeout(200);
  134 |     expect(await page.locator(".dxw-sel").count()).toBeGreaterThan(0);
  135 |   });
  136 | 
  137 |   test("copy puts selection text on the clipboard", async ({ page }) => {
  138 |     await load(page);
  139 |     const el = span(page, "exercises");
  140 |     const box = (await el.boundingBox())!;
  141 |     await page.mouse.dblclick(box.x + box.width / 2, box.y + box.height / 2);
  142 |     await page.waitForTimeout(200);
  143 |     await page.keyboard.press(`${MOD}+c`);
  144 |     await page.waitForTimeout(200);
  145 |     const clip = await page.evaluate(() => navigator.clipboard.readText());
  146 |     expect(clip).toContain("exercises");
  147 |   });
  148 | });
  149 | 
  150 | test.describe("tables", () => {
  151 |   test("column grip drag moves the boundary by the drag distance", async ({ page }) => {
  152 |     await load(page);
  153 |     const grips = page.locator("[data-dxw-grip]");
  154 |     let grip = grips.first();
  155 |     for (let i = 0; i < (await grips.count()); i++) {
  156 |       if ((await grips.nth(i).evaluate((el) => el.style.cursor)) === "col-resize") {
  157 |         grip = grips.nth(i);
  158 |         break;
  159 |       }
  160 |     }
  161 |     const gb = (await grip.boundingBox())!;
  162 |     const before = (await span(page, "Status").boundingBox())!.x;
  163 |     // Drag the Feature/Status boundary LEFT: autofit keeps these columns
  164 |     // near their content width, so growing into Status would starve it.
  165 |     await page.mouse.move(gb.x + 3, gb.y + 20);
  166 |     await page.mouse.down();
  167 |     await page.mouse.move(gb.x + 3 - 40, gb.y + 20, { steps: 4 });
  168 |     await page.mouse.up();
  169 |     await page.waitForTimeout(400);
  170 |     const after = (await span(page, "Status").boundingBox())!.x;
  171 |     expect(Math.abs(before - after - 40)).toBeLessThanOrEqual(3);
  172 |   });
  173 | 
  174 |   test("row grip drag grows the row", async ({ page }) => {
  175 |     await load(page);
  176 |     const rowGrip = page.locator("[data-dxw-grip]").filter({ has: page.locator(":scope") }).nth(0);
  177 |     // pick the first row-resize grip specifically
  178 |     const grips = page.locator("[data-dxw-grip]");
  179 |     const count = await grips.count();
  180 |     let target = null;
  181 |     for (let i = 0; i < count; i++) {
  182 |       const cursor = await grips.nth(i).evaluate((el) => el.style.cursor);
  183 |       if (cursor === "row-resize") {
  184 |         target = grips.nth(i);
  185 |         break;
  186 |       }
  187 |     }
  188 |     expect(target).not.toBeNull();
  189 |     const gb = (await target!.boundingBox())!;
  190 |     const before = (await span(page, "Pagination").boundingBox())!.y;
  191 |     await page.mouse.move(gb.x + 40, gb.y + 3);
  192 |     await page.mouse.down();
  193 |     await page.mouse.move(gb.x + 40, gb.y + 33, { steps: 4 });
  194 |     await page.mouse.up();
  195 |     await page.waitForTimeout(400);
  196 |     const after = (await span(page, "Pagination").boundingBox())!.y;
  197 |     expect(after - before).toBeGreaterThan(20);
  198 |     void rowGrip;
  199 |   });
  200 | 
  201 |   test("grid picker inserts a table at the caret", async ({ page }) => {
  202 |     await load(page);
  203 |     await clickText(page, "laborum.");
  204 |     const edgesBefore = await page.evaluate(
  205 |       () => [...document.querySelectorAll(".dxw-page div")].filter((d) => (d as HTMLElement).style.borderTop || (d as HTMLElement).style.borderLeft).length,
  206 |     );
  207 |     await page.locator('button[data-tab="insert"]').click();
  208 |     await page.waitForTimeout(100);
  209 |     await page.locator("button[title='Table']").click();
  210 |     await page.locator("div", { hasText: /^Insert table$/ }).first().waitFor();
  211 |     // click the 2x3 cell (row 2, col 3) in the 10-col grid
  212 |     const cells = page.locator("div").filter({ hasText: /^$/ });
  213 |     // simpler: use the grid cells by size
  214 |     const gridCells = page.locator("div[style*='width: 16px'][style*='height: 16px']");
  215 |     await gridCells.nth(12).click();
  216 |     await page.waitForTimeout(400);
  217 |     const edgesAfter = await page.evaluate(
  218 |       () => [...document.querySelectorAll(".dxw-page div")].filter((d) => (d as HTMLElement).style.borderTop || (d as HTMLElement).style.borderLeft).length,
  219 |     );
  220 |     expect(edgesAfter).toBeGreaterThan(edgesBefore);
  221 |     void cells;
  222 |   });
  223 | });
  224 | 
  225 | test.describe("headers and footers", () => {
  226 |   test("single click is gated; double-click enters with chrome; body double-click exits", async ({ page }) => {
  227 |     await load(page);
  228 |     const hdr = span(page, "Fidelity");
  229 |     const hb = (await hdr.boundingBox())!;
  230 |     await page.mouse.click(hb.x + 4, hb.y + 5);
  231 |     await page.waitForTimeout(150);
> 232 |     expect(await caretVisible(page)).toBe(false);
      |                                      ^ Error: expect(received).toBe(expected) // Object.is equality
  233 | 
  234 |     await page.mouse.dblclick(hb.x + 4, hb.y + 5);
  235 |     await page.waitForTimeout(250);
  236 |     expect(await caretVisible(page)).toBe(true);
  237 |     expect(await page.locator(".dxw-hf-marker").count()).toBeGreaterThan(0);
  238 |     const bodyOpacity = await span(page, "Lorem").evaluate((el) => getComputedStyle(el).opacity);
  239 |     expect(parseFloat(bodyOpacity)).toBeLessThan(1);
  240 | 
  241 |     // The dimmed body is inert: single clicks stay in header/footer mode.
  242 |     await clickText(page, "Lorem", "start");
  243 |     expect(await page.locator(".dxw-hf-marker").count()).toBeGreaterThan(0);
  244 | 
  245 |     // Double-click returns to body editing.
  246 |     const lb = (await span(page, "Lorem").boundingBox())!;
  247 |     await page.mouse.dblclick(lb.x + 1, lb.y + lb.height / 2);
  248 |     await page.waitForTimeout(250);
  249 |     expect(await page.locator(".dxw-hf-marker").count()).toBe(0);
  250 |   });
  251 | 
  252 |   test("body cannot be edited while header/footer mode is active", async ({ page }) => {
  253 |     await load(page);
  254 |     const hdr = span(page, "Fidelity");
  255 |     const hb = (await hdr.boundingBox())!;
  256 |     await page.mouse.dblclick(hb.x + 4, hb.y + 5);
  257 |     await page.waitForTimeout(250);
  258 | 
  259 |     // Drag-select over body text must not create a selection or edit it.
  260 |     const body = span(page, "Lorem");
  261 |     const bb = (await body.boundingBox())!;
  262 |     await page.mouse.move(bb.x + 2, bb.y + bb.height / 2);
  263 |     await page.mouse.down();
  264 |     await page.mouse.move(bb.x + 80, bb.y + bb.height / 2, { steps: 8 });
  265 |     await page.mouse.up();
  266 |     await page.waitForTimeout(150);
  267 |     expect(await page.locator(".dxw-sel").count()).toBe(0);
  268 |     await page.keyboard.type("ZAP");
  269 |     await page.waitForTimeout(300);
  270 |     expect(await span(page, "Lorem").count()).toBe(1);
  271 |     expect(await page.locator(".dxw-hf-marker").count()).toBeGreaterThan(0);
  272 |   });
  273 | 
  274 |   test("caret lands on the page whose header was double-clicked", async ({ page }) => {
  275 |     await load(page);
  276 |     const hdr2 = page.locator('.dxw-page[data-page="2"] span[data-dxw-hf]').first();
  277 |     await hdr2.scrollIntoViewIfNeeded();
  278 |     await page.waitForTimeout(200);
  279 |     const hb = (await hdr2.boundingBox())!;
  280 |     await page.mouse.dblclick(hb.x + 4, hb.y + hb.height / 2);
  281 |     await page.waitForTimeout(250);
  282 |     const caretPage = await page.evaluate(() => {
  283 |       const d = [...document.querySelectorAll("div")].find(
  284 |         (d) => d.style.width === "1.5px" && d.style.pointerEvents === "none" && d.style.display === "block",
  285 |       );
  286 |       return d?.closest(".dxw-page")?.getAttribute("data-page") ?? null;
  287 |     });
  288 |     expect(caretPage).toBe("2");
  289 |   });
  290 | });
  291 | 
  292 | test.describe("comments", () => {
  293 |   test("renders margin balloons with author/text, highlights ranges, hover links both ways", async ({ page }) => {
  294 |     await load(page);
  295 |     await expect(page.locator(".dxw-comment-card")).toHaveCount(2);
  296 |     const first = page.locator(".dxw-comment-card").first();
  297 |     await expect(first).toContainText("Ada Reviewer");
  298 |     await expect(first).toContainText("brand blue");
  299 |     expect(await page.locator("span[data-dxw-comment]").count()).toBeGreaterThan(0);
  300 | 
  301 |     // The balloon sits in the rail right of the page, near its anchor line.
  302 |     const pageBox = (await page.locator(".dxw-page").first().boundingBox())!;
  303 |     const cardBox = (await first.boundingBox())!;
  304 |     expect(cardBox.x).toBeGreaterThan(pageBox.x + pageBox.width);
  305 |     const anchor = (await page.locator("span[data-dxw-comment]").first().boundingBox())!;
  306 |     expect(Math.abs(cardBox.y - anchor.y)).toBeLessThan(60);
  307 | 
  308 |     // Hovering the commented text lights up its balloon.
  309 |     await page.locator("span[data-dxw-comment]").first().hover();
  310 |     await page.waitForTimeout(150);
  311 |     expect(await first.evaluate((el) => el.classList.contains("dxw-hot"))).toBe(true);
  312 |   });
  313 | 
  314 |   test("balloons stay anchored to the page when the window resizes", async ({ page }) => {
  315 |     await load(page);
  316 |     const gap = async () => {
  317 |       const pg = (await page.locator(".dxw-page").first().boundingBox())!;
  318 |       const cd = (await page.locator(".dxw-comment-card").first().boundingBox())!;
  319 |       return cd.x - (pg.x + pg.width);
  320 |     };
  321 |     const before = await gap();
  322 |     await page.setViewportSize({ width: 1150, height: 1000 });
  323 |     await page.waitForTimeout(300);
  324 |     expect(Math.abs((await gap()) - before)).toBeLessThan(2);
  325 |     await page.setViewportSize({ width: 1700, height: 1000 });
  326 |     await page.waitForTimeout(300);
  327 |     expect(Math.abs((await gap()) - before)).toBeLessThan(2);
  328 |   });
  329 | 
  330 |   test("replies nest in the parent balloon and round-trip through undo", async ({ page }) => {
  331 |     await load(page);
  332 |     const input = page.locator(".dxw-comment-card").first().locator(".dxw-comment-reply-input");
```