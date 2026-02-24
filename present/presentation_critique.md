# Presentation Critique: presentation.qmd

### 1. Title Assessment (Critical Flaw)
**Every single slide title in this presentation is a passive topic or noun phrase, rather than an active assertion.** This is a major design flaw. Passive titles force the audience to figure out the point of the slide themselves. 

**What is wrong & Why:** Titles like "The Expectation Trap" or "Strategic Alignment" tell the audience *what* you are talking about, but not *what they should conclude*. 
**How to fix it:** Rewrite every title to be a full, declarative sentence that states the core slide message. 

*   **Slide 1:** `Predicting Individual Purchasing Behavior` → **`Hurdle Models Predict Individual Purchasing Behavior More Accurately`**
*   **Slide 2:** `The Anatomy of the Data` → **`Full Distributions Obscure True Purchasing Behavior`**
*   **Slide 3:** `The Expectation Trap` → **`Blended Averages Predict Volumes That Customers Will Never Buy`**
*   **Slide 4:** `The Classification Shift` → **`Blended Models Falsely Flag Churn as Account Decline`**
*   **Slide 5:** `The Illusion of Decline` → **`Active Buyers Are Healthy; Our Primary Focus Must Be Churn Risk`**
*   **Slide 6:** `Portfolio Segmentation: The Two-Metric Advantage` → **`Two-Metric Segmentation Directs Sales Effort Where It Matters Most`**
*   **Slide 7:** `Strategic Alignment` → **`Align the Hurdle Model with Individual Sales Execution`**

### 2. Cognitive Load Assessment
Overall, the cognitive load of the text is managed well by keeping the heavy explanations in the speaker `::: {.notes}` and using concise bullet points/quotes on the slide canvas. However, there are a few visual overload risks:

*   **Slide 5 (The Illusion of Decline):** You are using `patchwork` to stack two complex density charts vertically.
    *   **What is wrong:** Showing two charts at the same time splits the audience's attention. They will try to spot the differences between the top and bottom charts while you are talking, rather than listening to you.
    *   **How to fix it:** Separate these into two distinct slides or use slide animations/fragments. Show "The Illusion" on its own slide first, make your point about the fake decline, and then progress to the next slide to reveal "The Reality". 
*   **Slide 4 & 5 Plot Legends/Annotations:** 
    *   **What is wrong:** You are using manual text annotations inside the plot canvas (e.g., `percent(prop_churn)`). If the underlying data changes, your text boxes might overlap with the curves or axes.
    *   **How to fix it:** The visual distinction using your custom categorical colors (`my_colors`) is strong enough. Rely on the colors and the verbal notes; prune the exact percentage annotations from the chart area to reduce visual clutter, or ensure they are placed with dynamic repulsive text (like `ggrepel`).
*   **Slide 7 (Strategic Alignment - Table):** 
    *   **What is wrong:** Tables inherently increase cognitive load because they require reading rather than viewing. The audience will read ahead.
    *   **How to fix it:** If you keep the table, use bolding strictly for the most critical words, not entire sentences in the right column. Prune the text to the absolute minimum required to make the comparison. 

### 3. Core Message Visibility
The core message—that separating conversion risk from order volume prevents sales reps from making strategic errors—is successfully represented in the data visualizations (specifically the isolated spike vs. the density curve). 

However, because the slide titles are currently passive (as identified in point #1), the core message is currently buried in the charts and speaker notes. Once the titles are converted to active assertions, the core message of each slide will be immediately obvious.
