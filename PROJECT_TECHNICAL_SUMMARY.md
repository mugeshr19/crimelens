# CrimeLens — Technical Summary: Algorithms, Formulas, Frameworks, Comparison & Data Generation

---

## 1. Algorithms & Formulas Used

All core intelligence algorithms are implemented in TypeScript under `CrimeLens_Full_Project/functions/` and `CrimeLens_Full_Project/server/`.

### 1.1 ST-DBSCAN (Spatio-Temporal Hotspot Clustering)
- **File:** `CrimeLens_Full_Project/functions/incidentCluster.ts` (`runSpatiotemporalDBSCAN`)
- **Purpose:** Groups crime incidents within a spatial radius ($\varepsilon_1 = 600\text{ m}$) and temporal window (`dateRangeDays`, `timeBand`) into localized risk hotspots.
- **Formulas:**
  - **Local tangent-plane distance** around anchor $(\phi_a, \lambda_a)$:
    $$\Delta\text{lat} = (\phi_i - \phi_a) \cdot 111000, \quad \Delta\text{lng} = (\lambda_i - \lambda_a) \cdot 111000 \cdot \cos\!\left(\frac{\phi_a \pi}{180}\right)$$
    $$d = \sqrt{\Delta\text{lat}^2 + \Delta\text{lng}^2} \le 600\text{ m}$$
  - **Intensity Score (0–100):**
    $$\text{intensityScore} = \min\!\left(98,\; \max\!\left(35,\; \text{round}\!\left(\frac{\text{count}}{40} \times 100\right)\right)\right)$$
  - **Emerging Trend Flag & Surge %:**
    $$\text{isEmergingTrend} = \left(\text{count} > 15 \times \frac{\text{dateRangeDays}}{30} \times 1.3\right), \quad \text{surgePct} = \text{round}\!\left(\frac{\text{count} - 15}{15} \times 100\right)$$

### 1.2 Station Anomaly Z-Score
- **File:** `CrimeLens_Full_Project/functions/incidentCluster.ts` (`computeStationAnomaly`)
- **Purpose:** Detects sudden weekly crime surges or drops per police station against its prior 11-week baseline.
- **Formula:**
  $$z = \frac{x_{\text{current}} - \mu}{\sigma}$$
  where $z \ge 2$ flags `SURGE`, $z \le -2$ flags `DROP`, otherwise `NORMAL`.

### 1.3 Louvain Modularity Community Detection & Bipartite Projection
- **File:** `CrimeLens_Full_Project/functions/graphAnalytics.ts` (`louvainCommunities`, `computeProjectionEdges`, `degreeMap`)
- **Purpose:** Detects organized criminal syndicates/gangs, identifies "kingpins" (highest live degree centrality per community), and infers hidden links between suspects sharing a vehicle, bank account, or location.
- **Formula (Modularity Gain):**
  $$\Delta Q = w_{\text{in}} - \frac{\Sigma_{\text{tot}} \cdot k_i}{2m}$$
  where $w_{\text{in}}$ is the sum of edge weights from node $i$ into community $C$, $\Sigma_{\text{tot}}$ is the total weighted degree of community $C$, $k_i$ is the weighted degree of node $i$, and $2m$ is the sum of all edge weights in the graph.

### 1.4 Holt's Linear Trend (Double Exponential Smoothing)
- **File:** `CrimeLens_Full_Project/functions/offlineAnalytics.ts` (`holtLinearForecast`)
- **Purpose:** Forecasts weekly crime incident counts 4 weeks ahead with 80% prediction intervals (used in Crime Trends & Predictive Patrol Planner).
- **Formulas ($\alpha = 0.5,\; \beta = 0.3$):**
  $$L_t = \alpha y_t + (1 - \alpha)(L_{t-1} + T_{t-1})$$
  $$T_t = \beta(L_t - L_{t-1}) + (1 - \beta)T_{t-1}$$
  $$\hat{y}_{t+h} = \max\!\left(0,\; \text{round}(L_t + h \cdot T_t)\right)$$
  $$\text{RMSE} = \sqrt{\frac{1}{n}\sum (y_i - \hat{y}_i)^2}, \quad \text{80\% band}_h = \text{round}\!\left(1.28 \cdot \text{RMSE} \cdot \sqrt{h}\right)$$

### 1.5 TF-IDF + Cosine Similarity & Behavioral MO Clustering
- **Files:** `CrimeLens_Full_Project/functions/offlineAnalytics.ts` (`moSemanticSearch`), `CrimeLens_Full_Project/functions/analystFeatures.ts` (`behavioralSignatures`, `predictiveSuspectRanking`)
- **Purpose:** Semantic modus operandi (MO) search, similar-case matching, grounded RAG retrieval for the AI Copilot, agglomerative behavioral signature clustering, and predictive suspect ranking.
- **Formulas:**
  - **Term Frequency & Smoothed IDF:**
    $$\text{tf}(t, d) = \frac{f_{t,d}}{|d|}, \quad \text{idf}(t) = \ln\!\left(\frac{N + 1}{\text{df}(t) + 1}\right) + 1$$
  - **Cosine Similarity:**
    $$\cos(\mathbf{q}, \mathbf{d}) = \frac{\mathbf{q} \cdot \mathbf{d}}{\|\mathbf{q}\| \|\mathbf{d}\|}$$
  - **Behavioral Token Cosine (threshold $\ge 0.18$):**
    $$\text{sim}(A, B) = \frac{|A \cap B|}{\sqrt{|A||B|}}$$
  - **Predictive Suspect Lead Score:**
    $$\text{score} = 0.6 \cdot \frac{\text{degree}}{\max(\text{degree})} + 0.4 \cdot \text{moSim}$$

### 1.6 Haversine Great-Circle Distance & Nearest-Unit Dispatch
- **File:** `CrimeLens_Full_Project/functions/fleetTelemetry.ts` (`haversineM`, `dispatchNearest`)
- **Purpose:** Computes real-world geodesic distance ($R = 6,371,000\text{ m}$) from on-duty Akka Patrol units to Panic/SOS or hotspot coordinates and estimates dispatch ETA.
- **Formulas:**
  $$s = \sin^2\!\left(\frac{\Delta\phi}{2}\right) + \cos(\phi_1)\cos(\phi_2)\sin^2\!\left(\frac{\Delta\lambda}{2}\right), \quad d = 2R \arcsin(\sqrt{s})$$
  $$\text{ETA (min)} = \max\!\left(1,\; \text{round}\!\left(\frac{d}{400}\right)\right) \quad (\text{assuming } \approx 24\text{ km/h urban speed})$$

### 1.7 Edge Computer Vision (COCO-SSD)
- **File:** `CrimeLens_Full_Project/web/src/pages/CameraIntelligence.tsx`
- **Purpose:** Client-side real-time object detection (`@tensorflow/tfjs` + `@tensorflow-models/coco-ssd` `ssdlite_mobilenet_v2`) on CCTV/camera streams and ANPR vehicle-match alerts without streaming raw video externally.

### 1.8 SHA-256 Cryptographic Hash Chaining
- **Files:** `CrimeLens_Full_Project/functions/governanceAndCopilot.ts` (`verifyAuditLedgerIntegrity`), `CrimeLens_Full_Project/server/police/store.ts` (`appendAudit`, `verifyAuditChain`)
- **Purpose:** Append-only, tamper-evident audit ledger for governance actions and evidence chain-of-custody.
- **Formula:**
  $$H_i = \text{SHA256}(\text{payload}_i \parallel H_{i-1})$$

---

## 2. Frameworks & Libraries Used

| Layer | Frameworks / Libraries | Purpose |
|---|---|---|
| **Mobile Field App** | **Expo SDK 54**, **React Native 0.81**, **Expo Router v6**, `expo-location`, `expo-sensors`, `expo-sqlite` | Cross-platform offline-first officer app (beat check-ins, evidence capture, one-tap Panic/SOS). |
| **Web Command Console** | **React 18**, **Vite 5**, **TailwindCSS 3.4**, **React Router v6**, **Framer Motion** | Interactive role-based dashboards for SPs, Analysts, Investigating Officers, and Super Admins. |
| **Visualization & Maps** | **Leaflet / react-leaflet**, **react-force-graph-2d**, **Recharts** | Hotspot cluster maps, force-directed gang network graphs, time-series trend & forecast charts. |
| **AI / ML** | **TensorFlow.js** (`@tensorflow/tfjs`, `@tensorflow-models/coco-ssd`), **Groq API** (`llama-3.1-8b-instant`) | On-device CCTV object detection + grounded RAG intelligence copilot. |
| **Backend & Real-Time** | **Node.js 20**, **Express 5**, **TypeScript**, **ws** (WebSockets), **jsonwebtoken**, **bcryptjs**, **@simplewebauthn/server** | REST API, live SOS/crime alert push bus, RBAC, JWT & biometric WebAuthn authentication. |
| **Database & Deployment** | **Drizzle ORM**, **PostgreSQL (`pg`)**, **Zoho Catalyst (`ZCQL`)**, **esbuild** | Schema management, serverless/production bundling. |

---

## 3. Difference Between Existing Systems and CrimeLens

| Dimension | Existing Systems (Legacy CCTNS / Commercial GIS / ML Scripts) | CrimeLens |
|---|---|---|
| **Data Plane** | Siloed per-station SQL FIR entry; cross-district sharing is manual | Unified ingestion of 201,733 KSP records across 37 districts + automated cross-station case linkage |
| **Hotspot Analysis** | Static tabular queries or opaque kernel density heatmaps | Dual-parameter **ST-DBSCAN** (spatial + temporal) with emerging-trend surge flags |
| **Syndicate Detection** | Manual name matching or basic adjacency lists | Automated **Louvain modularity** community detection + bipartite entity projection + kingpin centrality |
| **Forecasting** | Retrospective bar charts or linear extrapolation | **Holt's double-exponential smoothing** with RMSE & 80% prediction bands |
| **Field Deployment** | Desktop-only station terminals | Offline-capable **Expo mobile app** with live GPS telemetry & Haversine SOS dispatch |
| **Audit & Governance** | Mutable SQL logs; black-box AI scores | **SHA-256 hash-chained** audit ledger + plain-language rationale & feature weights |

---

## 4. How CrimeLens Is Better Than Existing Systems

1. **Breaks Jurisdiction Silos:** Cross-station case linkage (`server/police/store.ts`) and Louvain network graphs automatically connect suspects operating across different police stations via shared MOs, vehicles, or bank accounts.
2. **Proactive Instead of Reactive:** Combines ST-DBSCAN hotspots, Holt's 4-week forecast, and time-slot risk matrices into deployable patrol plans before crime spikes occur.
3. **Courtroom-Ready Governance & Explainability:** Every AI recommendation is paired with human-readable reasons (matched MO terms, centrality score, distance) and every supervisory action is sealed in a verifiable SHA-256 hash chain.
4. **End-to-End Field-to-Command Loop:** Connects ground officers (mobile SOS, beat checklists, QR evidence custody) directly to the command center via WebSockets and Haversine nearest-unit dispatch.
5. **Privacy-Preserving Edge AI:** Runs TensorFlow.js COCO-SSD directly in the browser on camera feeds without streaming raw video to external servers.

---

## 5. How Data Generation & Ingestion Is Done

CrimeLens uses a **hybrid two-tier data pipeline** combining real aggregate Karnataka crime data with synthetic/seeded operational records:

1. **Real KSP Corpus Ingestion** (`server/dataset-loader.ts`, `functions/ingestRealDataset.ts`)
   - Reads **12 CSV files** from `CrimeLens_Full_Project/dataset/` (including `ka-district-wise-2025.csv`, `ka-ipc-crimes-2025.csv`, `ka-sll-crimes-2025.csv`, and `crime_review_for_the_month_of_december_2025_9.csv`), totaling **201,733 records across 37 districts** for real district-level totals, IPC/SLL breakdowns, and hotspot rankings.
2. **Synthetic Spatio-Temporal Incident Generation** (`functions/incidentCluster.ts` — `seedIncidentDatabase()`)
   - Because public KSP CSVs are district-level aggregates rather than individual GPS-tagged FIRs, `seedIncidentDatabase()` generates **2,200 synthetic incidents**:
     - **1,500 clustered incidents** distributed around **10 real Karnataka `HOTSPOT_ANCHORS`** (e.g., Peenya, Koramangala, Hebbal, Gokulam, Hampankatta, Hubballi, Belagavi, Shivamogga) with ~300 m random lat/lng offsets ($\pm 0.003^\circ$), timestamps spanning the last 90 days, and a **65% nighttime bias (23:00–03:00)**.
     - **700 background scattered incidents** uniformly distributed across 8 Karnataka districts.
3. **Seeded Operational & Anonymized Entity Data** (`server/police/store.ts` — `seed()`, `functions/ingestRealDataset.ts`)
   - Seeds anonymized suspects, gangs, vehicles, bank accounts, stations, officers, beat checkpoints, and initial FIRs/cases in memory so all CRUD, cross-station linkage, and graph analytics work out of the box.

