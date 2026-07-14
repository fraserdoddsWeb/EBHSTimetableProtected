# School Timetable App v3 working baseline optimiser

This version restores the last known working current-split baseline route, then tries to repair 3+ and 4+ split classes without using non-specialist teachers.

Build label shown in the app:

`teacher-first optimiser v7 working baseline repair`

Recommended first settings:

- Teacher-first strict specialist solve: ON
- Teacher-first: use current split allocation from JSON: OFF
- No emergency non-specialist teaching: ON
- Teacher-first allocation attempts: 3
- Teacher-first timing attempts per allocation: 20
- Stop after this many seconds: 240

The app now starts from a working timetable and only replaces it if an improved working timetable is found.
