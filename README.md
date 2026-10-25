# Astana travel-time estimates

RMT Assignment 3, group SE-2424. Yerdaulet Akhanov, Syrym Shadiyarbek and Akylbek Sabyrzhan.

Automated API requests sampled 20 directed routes between 10 public locations in Astana. The comparison uses Google Routes with `TRAFFIC_AWARE_OPTIMAL` and openrouteservice (ORS) without live traffic. These are provider estimates; the study did not observe journeys on the road.

This dataset contains 404 ORS observations: 60 pilot, 284 scheduled, 20 exploratory and 40 follow-up records. ORS local dates: 2026-10-24, 2026-10-25. The first request started at 2026-10-24T02:37:23.929322+00:00; the latest archived study timestamp is 2026-10-25T04:29:44.181817+00:00 (2026-10-25T09:29:44.181817+05:00 in Astana). Window status describes the evidence available at that timestamp.

Google has 60 request-status records and 3 aggregate comparison batches, covering 60 valid route pairs on local dates 2026-10-24. The follow-up plan has 4 windows, of which 2 have attempted every planned route. The exploratory batch has an empty window_id and remains separate from the planned comparisons.

| File | Contents |
| --- | --- |
| `routes.csv` | 20 directed routes, coordinates and OpenStreetMap source links |
| `ors_observations.csv` | 404 original ORS measurement records |
| `ors_processed.csv` | The same 404 records with local period/date and quality checks |
| `google_request_status.csv` | 60 Google request outcomes and timestamps |
| `google_ors_summary.csv` | 3 Google-ORS batches with all recorded aggregate statistics |
| `collection_windows.csv` | 4 planned follow-up windows and recorded completion status |
| `selected_endpoints.csv` | 10 endpoint sources and selection rules |
| `ors_pilot_routes.geojson` | 20 route geometries from the first ORS pilot cycle |
| `data_dictionary.csv` | Definitions and units for every CSV column and GeoJSON field |

Raw ORS measurement values, timestamps and collection identifiers retain their original text. The compact raw table omits file-path fields for response and geometry files that are not part of this release. Processed records add derived fields without replacing any raw measurement. 404 of 404 ORS rows pass the published checks. Blank optional values mean unavailable, never zero. All CSV files use UTF-8 and comma delimiters.

Astana uses UTC+05:00 for these study dates. Peak is 17:30 to 19:30; off-peak is 10:00 to 12:00 or 22:00 to 24:00. Each interval includes its start and excludes its end. The processed table checks request success, positive finite duration and distance, speed between 2 and 200 km/h, timestamps and duplicate identities. It retains excluded rows with reasons. Google follow-up pairs required successful responses within their planned window and request starts no more than 60 seconds apart.

Google individual durations, distances and geometry are absent because the study retained only aggregate statistics and request status. The Google aggregates therefore cannot be recalculated from individual archived responses. Signed differences are Google minus ORS. Repeated routes and batches share locations; request counts are not independent sample counts. Provider differences describe disagreement and do not establish accuracy or isolate the effect of traffic.

Route coordinates and endpoint links identify existing OpenStreetMap road nodes near public landmarks. They do not establish building entrances. `expected_length_m` is the spherical straight-line lower bound, while `distance_m` is the route length returned by ORS. The pilot GeoJSON preserves the original ORS coordinates; it does not represent Google routes or prove that two providers chose the same path.

ORS attribution: © openrouteservice by HeiGIT | Data from OpenStreetMap. See [openrouteservice](https://openrouteservice.org/), its [terms](https://account.heigit.org/info/tos), and [CC BY-SA 4.0](https://creativecommons.org/licenses/by-sa/4.0/). OpenStreetMap attribution: © OpenStreetMap contributors; see the [copyright page](https://www.openstreetmap.org/copyright) and [ODbL 1.0](https://opendatacommons.org/licenses/odbl/1-0/). Individual source URLs appear in the route and endpoint tables. Google source: [Routes API](https://developers.google.com/maps/documentation/routes/overview) and [traffic preferences](https://developers.google.com/maps/documentation/routes/config_trade_offs). Provider terms apply to the respective material; this repository does not assign a single license to all sources.
