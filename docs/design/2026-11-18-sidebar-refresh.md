# Sidebar visual refresh
Intent: quiet companion for reading, preserving existing page, AI guide, personal notes, library and settings flows.
Direction: Apple-style hierarchy and restraint, informed by apple-design (emilkowalski/skills) and frontend-design (anthropics/skills); system type keeps Chinese and Latin readable. Readwise informs the distinction between source passage and annotation.
Palette: paper #f6f7f9, surface #ffffff, ink #24262b, secondary #616773, line #e3e6ec, accent #365fc3. Dark: #1c1d21, #24262b, #eceef3, #aeb3be, #363941, #a5bfff.
Type: platform system; article heading 15px/1.45 weight600, passage 14px/1.6 weight550, note 13px/1.75 weight400, controls13px.
Layout: compact article header → workdesk/library actions → search → stable tab navigation → unboxed reading content. Settings use one disclosure per group.
Alternatives: retain current bordered dashboard (too much competing chrome); dramatic glass (risk of text interference); choose quiet reading utility.
Implementation: consolidate sidebar CSS, add coherent outline toolbar icons, strengthen selected navigation, soften failure notice while preserving it, visually separate source from thought. No extra dependencies or network assets.
Verification: all existing Node checks; browser fixtures at 320/400/640 and light/dark; inspect screenshots with realistic passages, notes and settings.
