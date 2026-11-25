import matplotlib.pyplot as plt
import matplotlib.patches as patches
from matplotlib.patches import Rectangle, Circle, Ellipse, FancyBboxPatch
import numpy as np

def draw_cylinder(ax, x, y, w, h, label, fontsize=6.8):
    # Cylinder top ellipse
    top_h = h * 0.22
    # Cylinder body
    rect = Rectangle((x, y), w, h - top_h/2, facecolor='#FFFFFF', edgecolor='#000000', linewidth=1.2, zorder=2)
    ax.add_patch(rect)
    
    # Cylinder bottom curve
    bottom_ellipse = Ellipse((x + w/2, y), w, top_h, facecolor='#FFFFFF', edgecolor='#000000', linewidth=1.2, zorder=3)
    ax.add_patch(bottom_ellipse)
    
    # Redraw rect bottom without line
    rect_inner = Rectangle((x + 0.1, y + 0.1), w - 0.2, h - top_h/2 - 0.2, facecolor='#FFFFFF', edgecolor='none', zorder=4)
    ax.add_patch(rect_inner)
    
    # Cylinder top curve
    top_ellipse = Ellipse((x + w/2, y + h - top_h/2), w, top_h, facecolor='#FFFFFF', edgecolor='#000000', linewidth=1.2, zorder=5)
    ax.add_patch(top_ellipse)
    
    # Label
    ax.text(x + w/2, y + (h - top_h/2)/2, label, fontsize=fontsize, ha='center', va='center',
            color='#000000', fontweight='bold', zorder=6, linespacing=1.2, fontfamily='DejaVu Sans')

def draw_stick_figure(ax, x, y, scale=1.0):
    # Head
    head = Circle((x, y + 3.8 * scale), 1.4 * scale, facecolor='#FFFFFF', edgecolor='#000000', linewidth=1.3, zorder=10)
    ax.add_patch(head)
    
    # Body
    ax.plot([x, x], [y + 2.4 * scale, y - 1.2 * scale], color='#000000', linewidth=1.3, zorder=10)
    
    # Arms
    ax.plot([x - 2.2 * scale, x, x + 2.2 * scale], [y + 0.6 * scale, y + 1.8 * scale, y + 0.6 * scale], color='#000000', linewidth=1.3, zorder=10)
    
    # Legs
    ax.plot([x - 1.6 * scale, x, x + 1.6 * scale], [y - 3.4 * scale, y - 1.2 * scale, y - 3.4 * scale], color='#000000', linewidth=1.3, zorder=10)

def create_diagram():
    fig, ax = plt.subplots(figsize=(9, 16.5), dpi=300)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 192)
    ax.axis('off')
    fig.patch.set_facecolor('#FFFFFF')

    # Main block bounds
    BX = 18.0
    BW = 64.0
    
    # Helper to draw a layer box
    def draw_layer_box(y_bottom, height, layer_num, layer_title):
        # Outer box
        rect = Rectangle((BX, y_bottom), BW, height, facecolor='#FFFFFF', edgecolor='#000000', linewidth=1.3, zorder=2)
        ax.add_patch(rect)
        
        # Header text
        header_text = f"{layer_num}   {layer_title}"
        ax.text(BX + 1.2, y_bottom + height - 1.5, header_text, fontsize=7.8, fontweight='bold',
                ha='left', va='top', color='#000000', fontfamily='DejaVu Sans')
        return (BX, y_bottom, BW, height)

    def draw_sub_boxes(x_start, y_start, total_w, total_h, rows_data):
        num_rows = len(rows_data)
        pad_x = 0.8
        pad_y = 0.7
        h_per_row = (total_h - (num_rows + 1) * pad_y) / num_rows
        
        for r_idx, row in enumerate(rows_data):
            num_cols = len(row)
            w_per_col = (total_w - (num_cols + 1) * pad_x) / num_cols
            curr_y = y_start + total_h - pad_y - (r_idx + 1) * h_per_row - r_idx * pad_y
            
            for c_idx, item in enumerate(row):
                curr_x = x_start + pad_x + c_idx * (w_per_col + pad_x)
                srect = Rectangle((curr_x, curr_y), w_per_col, h_per_row, facecolor='#FFFFFF', edgecolor='#000000', linewidth=0.9, zorder=3)
                ax.add_patch(srect)
                
                # Check if item is tuple (title, [bullets]) or simple string
                if isinstance(item, tuple):
                    title, bullets = item
                    ax.text(curr_x + w_per_col/2, curr_y + h_per_row - 1.6, title, fontsize=6.8, fontweight='bold',
                            ha='center', va='top', color='#000000', fontfamily='DejaVu Sans', linespacing=1.1)
                    bullet_y = curr_y + h_per_row - 4.2
                    for b in bullets:
                        ax.text(curr_x + 1.0, bullet_y, b, fontsize=5.8, ha='left', va='top',
                                color='#000000', fontfamily='DejaVu Sans')
                        bullet_y -= 2.0
                else:
                    ax.text(curr_x + w_per_col/2, curr_y + h_per_row/2, item, fontsize=6.4,
                            ha='center', va='center', color='#000000', fontfamily='DejaVu Sans', linespacing=1.15)

    def draw_down_arrow(y_top, y_bottom):
        ax.annotate('', xy=(BX + BW/2, y_bottom), xytext=(BX + BW/2, y_top),
                    arrowprops=dict(facecolor='#000000', edgecolor='#000000', width=1.0, headwidth=4.5, headlength=4.5),
                    zorder=8)

    # -------------------------------------------------------------
    # LAYER 1: DATA SOURCES
    # -------------------------------------------------------------
    L1_Y = 169.5
    L1_H = 17.5
    draw_layer_box(L1_Y, L1_H, "1", "DATA SOURCES")
    l1_data = [
        ["KSP FIR Corpus\n(201,733 records)", "IPC / SLL\nStatutory Codes", "CCTNS Real-time\nIncident Feeds", "Monthly Crime\nReview Extracts"],
        ["ANPR & CCTV\nTelemetry", "Akka Patrol Unit\nGPS Telemetry", "Modus Operandi\nPolice Diaries", "Emergency 112\nCAD Telemetry"]
    ]
    draw_sub_boxes(BX + 0.6, L1_Y + 0.6, BW - 1.2, L1_H - 3.2, l1_data)

    draw_down_arrow(L1_Y, L1_Y - 2.5)

    # -------------------------------------------------------------
    # LAYER 2: COLLECTION AND PREPROCESSING
    # -------------------------------------------------------------
    L2_Y = 149.0
    L2_H = 18.0
    draw_layer_box(L2_Y, L2_H, "2", "COLLECTION AND PREPROCESSING")
    l2_data = [
        ["API / Batch Ingest\n(12 CSV Files)", "Geospatial Coordinate\nValidation", "Temporal Window\nExtraction", "FIR Narrative\nText Extraction"],
        ["Cleaning &\nDeduplication", "Missing Coordinate\nHandling", "Normalization &\nStatutory Encoding", "Schema Validation\n(CaseMaster)"]
    ]
    draw_sub_boxes(BX + 0.6, L2_Y + 0.6, BW - 1.2, L2_H - 3.2, l2_data)

    draw_down_arrow(L2_Y, L2_Y - 2.5)

    # -------------------------------------------------------------
    # LAYER 3: CRIME INTELLIGENCE + MULTI-AGENT CORE
    # -------------------------------------------------------------
    L3_Y = 128.5
    L3_H = 18.0
    draw_layer_box(L3_Y, L3_H, "3", "CRIME INTELLIGENCE + MULTI-AGENT CORE")
    l3_data = [
        ["Spatio-Temporal\nHotspot Agent", "Syndicate Network\nGraph Agent", "Time-Series Crime\nForecast Agent", "Modus Operandi\nMatcher Agent"],
        ["Edge Computer\nVision Agent", "Akka Dispatch\nOptimizer Agent", "QuickML NL Search\n(English + Kannada)", "Cross-District\nCorrelator Agent"]
    ]
    draw_sub_boxes(BX + 0.6, L3_Y + 0.6, BW - 1.2, L3_H - 3.2, l3_data)

    # Database cylinder 1
    draw_cylinder(ax, 84.5, L3_Y + 2.0, 9.5, 14.0, "State Crime\nStore &\nMO Memory", fontsize=6.2)
    # Connector between L3 and Cylinder 1
    ax.plot([BX + BW, 84.5], [L3_Y + L3_H/2, L3_Y + L3_H/2], color='#000000', linewidth=1.1, zorder=2)

    draw_down_arrow(L3_Y, L3_Y - 2.5)

    # -------------------------------------------------------------
    # LAYER 4: FEATURE ENGINEERING
    # -------------------------------------------------------------
    L4_Y = 114.5
    L4_H = 11.5
    draw_layer_box(L4_Y, L4_H, "4", "FEATURE ENGINEERING")
    l4_data = [
        ["Spatial Coordinates &\nHaversine Distances", "Temporal Windowing\n(Delta t = 7 Days)", "Co-Offender Adjacency\n& Centrality Metrics", "TF-IDF Modus\nOperandi Vectors"]
    ]
    draw_sub_boxes(BX + 0.6, L4_Y + 0.6, BW - 1.2, L4_H - 3.2, l4_data)

    draw_down_arrow(L4_Y, L4_Y - 2.5)

    # -------------------------------------------------------------
    # LAYER 5: MODEL AND ANALYTICS LAYER
    # -------------------------------------------------------------
    L5_Y = 91.0
    L5_H = 21.0
    draw_layer_box(L5_Y, L5_H, "5", "MODEL AND ANALYTICS LAYER")
    l5_data = [
        [
            ("SPATIO-TEMPORAL\nCLUSTERING", ["• ST-DBSCAN", "• eps1 = 600 m", "• dt = 7 days", "• MinPts = 4"]),
            ("CRIMINAL NETWORK\nGRAPH", ["• Louvain Modularity", "• Q Optimization", "• Centrality Rank", "• Confirmed / Inferred"]),
            ("TIME-SERIES\nFORECASTING", ["• Holt's Double Exp.", "• Level (Lt) & Trend (Tt)", "• 4-Week Horizon", "• RMSE & MAE Eval"]),
            ("EDGE VISION & XAI\nMODELS", ["• MobileNet COCO-SSD", "• Client-side ANPR OCR", "• Cosine MO Clusters", "• Plain-text XAI Drawers"])
        ]
    ]
    draw_sub_boxes(BX + 0.6, L5_Y + 0.6, BW - 1.2, L5_H - 3.2, l5_data)

    # Database cylinder 2
    draw_cylinder(ax, 84.5, L5_Y + 3.5, 9.5, 14.0, "Model\nRepository &\nWeights", fontsize=6.2)
    # Connector between L5 and Cylinder 2
    ax.plot([BX + BW, 84.5], [L5_Y + L5_H/2, L5_Y + L5_H/2], color='#000000', linewidth=1.1, zorder=2)

    draw_down_arrow(L5_Y, L5_Y - 2.5)

    # -------------------------------------------------------------
    # LAYER 6: ANALYSIS OUTPUTS
    # -------------------------------------------------------------
    L6_Y = 77.0
    L6_H = 11.5
    draw_layer_box(L6_Y, L6_H, "6", "ANALYSIS OUTPUTS")
    l6_data = [
        ["Hotspot Clusters\n(Persistent vs Noise)", "Gang Community\nStructures & Kingpins", "Jurisdictional Trend\nForecasts (4 Weeks)", "Modus Operandi\nCrime Signatures"]
    ]
    draw_sub_boxes(BX + 0.6, L6_Y + 0.6, BW - 1.2, L6_H - 3.2, l6_data)

    draw_down_arrow(L6_Y, L6_Y - 2.5)

    # -------------------------------------------------------------
    # LAYER 7: DECISION AND APPROVAL (ORCHESTRATOR)
    # -------------------------------------------------------------
    L7_Y = 57.5
    L7_H = 17.0
    draw_layer_box(L7_Y, L7_H, "7", "DECISION AND APPROVAL (ORCHESTRATOR)")
    l7_data = [
        ["Tiered Decision\nPolicy (L1-L4)", "Human-in-the-Loop\nApproval Gate", "SHA-256 Tamper-Evident\nAudit Ledger", "Role-Based Access\nControl (4 Tiers)"]
    ]
    draw_sub_boxes(BX + 0.6, L7_Y + 0.6, BW - 1.2, L7_H - 3.2, l7_data)

    draw_down_arrow(L7_Y, L7_Y - 2.5)

    # -------------------------------------------------------------
    # LAYER 8: PRESENTATION (DUAL INTERFACES)
    # -------------------------------------------------------------
    L8_Y = 38.0
    L8_H = 17.0
    draw_layer_box(L8_Y, L8_H, "8", "PRESENTATION (DUAL INTERFACES)")
    l8_data = [
        ["State Command\nCenter Console", "Interactive Hotspot\n& Network Maps", "Predictive Patrol\nPlanner Dashboard", "Field-Officer Mobile\nApp (Expo SDK 54)"]
    ]
    draw_sub_boxes(BX + 0.6, L8_Y + 0.6, BW - 1.2, L8_H - 3.2, l8_data)

    draw_down_arrow(L8_Y, L8_Y - 2.5)

    # -------------------------------------------------------------
    # LAYER 9: OPERATIONAL OUTPUTS
    # -------------------------------------------------------------
    L9_Y = 18.5
    L9_H = 17.0
    draw_layer_box(L9_Y, L9_H, "9", "OPERATIONAL OUTPUTS")
    l9_data = [
        ["Risk-Weighted Patrol\nDeployments", "Nearest-Unit Akka\nSOS Dispatches", "Cross-District\nFIR Linkages", "Courtroom-Admissible\nAudit Trails"]
    ]
    draw_sub_boxes(BX + 0.6, L9_Y + 0.6, BW - 1.2, L9_H - 3.2, l9_data)

    # -------------------------------------------------------------
    # PERSONA / ACTOR ON THE LEFT (Security Analyst / Approver)
    # -------------------------------------------------------------
    PERSON_X = 9.5
    PERSON_Y = 181.0
    draw_stick_figure(ax, PERSON_X, PERSON_Y, scale=1.0)
    ax.text(PERSON_X, PERSON_Y - 5.5, "Supervisory Officer /\nCommand Approver", fontsize=6.8, fontweight='bold',
            ha='center', va='top', color='#000000', fontfamily='DejaVu Sans', linespacing=1.1)

    # Feedback Line from Layer 9 back to Approver
    FEED_X = 5.0
    # Line from Layer 9 left edge to FEED_X
    ax.plot([BX, FEED_X], [L9_Y + L9_H/2, L9_Y + L9_H/2], color='#000000', linewidth=1.1)
    # Line going up along FEED_X
    ax.plot([FEED_X, FEED_X], [L9_Y + L9_H/2, PERSON_Y - 1.0], color='#000000', linewidth=1.1)
    # Arrow to person
    ax.annotate('', xy=(PERSON_X - 2.5, PERSON_Y), xytext=(FEED_X, PERSON_Y - 1.0),
                arrowprops=dict(facecolor='#000000', edgecolor='#000000', width=0.9, headwidth=4.0, headlength=4.0))
    # "Feedback" vertical text
    ax.text(FEED_X - 1.2, (L9_Y + PERSON_Y) / 2, "Feedback", fontsize=6.8, rotation=90,
            ha='center', va='center', color='#000000', fontfamily='DejaVu Sans')

    # "approve / reject" arrow from FEED_X into Layer 7
    ax.annotate('', xy=(BX, L7_Y + L7_H/2), xytext=(FEED_X, L7_Y + L7_H/2),
                arrowprops=dict(facecolor='#000000', edgecolor='#000000', width=1.0, headwidth=4.5, headlength=4.5))
    ax.text((FEED_X + BX)/2, L7_Y + L7_H/2 + 1.2, "approve /\nreject", fontsize=6.2,
            ha='center', va='bottom', color='#000000', fontfamily='DejaVu Sans', linespacing=1.0)

    # -------------------------------------------------------------
    # RIGHT SIDE: OUTCOMES WRITTEN BACK (Dashed line)
    # -------------------------------------------------------------
    DASH_X = 96.5
    # From Layer 9 right edge to DASH_X
    ax.plot([BX + BW, DASH_X], [L9_Y + L9_H/2, L9_Y + L9_H/2], color='#000000', linewidth=1.1, linestyle='--')
    # Vertical dashed line up to Cylinder 1
    ax.plot([DASH_X, DASH_X], [L9_Y + L9_H/2, L3_Y + 4.0], color='#000000', linewidth=1.1, linestyle='--')
    # Arrow into Cylinder 1
    ax.annotate('', xy=(94.0, L3_Y + 4.0), xytext=(DASH_X, L3_Y + 4.0),
                arrowprops=dict(facecolor='#000000', edgecolor='#000000', width=0.9, headwidth=4.0, headlength=4.0, linestyle='--'))
    # Vertical text "outcomes written back"
    ax.text(DASH_X - 1.0, (L9_Y + L3_Y) / 2 + 5.0, "outcomes written back", fontsize=6.5, rotation=-90,
            ha='center', va='center', color='#000000', fontfamily='DejaVu Sans')

    plt.tight_layout()
    out_path = 'd:/projects/CrimeLens/CrimeLens_Full_Project/docs/screenshots/crimelens_architecture_diagram_academic.png'
    plt.savefig(out_path, bbox_inches='tight', dpi=300)
    print(f"Diagram successfully saved to {out_path}")

if __name__ == "__main__":
    create_diagram()
