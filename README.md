# Sewangan Main App — Android-first v2
API is preconfigured to the production Apps Script /exec URL supplied by the user.
Temporary National-level test member:
Phone: 9999999999
Password: national02
Designation: National Secretary
Member ID: SCT/2026/TEST/00001
Valid until: 2026-12-31

Performance:
- one bootstrap request after login
- local cached bootstrap for instant reopen
- Apps Script CacheService for sheet reads
- module data loaded only when opened
- no repeated 5-second reloads
- compact mobile UI modeled after the Sursand Connect card/tile approach
