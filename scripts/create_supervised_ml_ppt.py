#!/usr/bin/env python3
"""Generate a PowerPoint presentation explaining Supervised Machine Learning."""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# Color palette
PRIMARY = RGBColor(0x1A, 0x56, 0xDB)      # Deep blue
SECONDARY = RGBColor(0x0E, 0x9F, 0x6E)    # Teal green
ACCENT = RGBColor(0xF5, 0x9E, 0x0B)       # Amber
DARK = RGBColor(0x1F, 0x29, 0x37)         # Near black
LIGHT_BG = RGBColor(0xF3, 0xF4, 0xF6)     # Light gray
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
MUTED = RGBColor(0x6B, 0x72, 0x80)


def set_slide_bg(slide, color):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_title_bar(slide, title_text, subtitle_text=None):
    """Add a colored header bar with title."""
    bar = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(10), Inches(1.2)
    )
    bar.fill.solid()
    bar.fill.fore_color.rgb = PRIMARY
    bar.line.fill.background()

    title_box = slide.shapes.add_textbox(Inches(0.6), Inches(0.25), Inches(8.8), Inches(0.7))
    tf = title_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = title_text
    p.font.size = Pt(32)
    p.font.bold = True
    p.font.color.rgb = WHITE
    p.font.name = "Calibri"

    if subtitle_text:
        sub_box = slide.shapes.add_textbox(Inches(0.6), Inches(1.35), Inches(8.8), Inches(0.4))
        stf = sub_box.text_frame
        sp = stf.paragraphs[0]
        sp.text = subtitle_text
        sp.font.size = Pt(14)
        sp.font.color.rgb = MUTED
        sp.font.name = "Calibri"


def add_bullets(slide, items, left=0.6, top=1.8, width=8.8, height=5.0, font_size=20, level_indent=0.3):
    """Add bullet points to a slide."""
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = box.text_frame
    tf.word_wrap = True

    for i, item in enumerate(items):
        if isinstance(item, tuple):
            text, level = item
        else:
            text, level = item, 0

        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = text
        p.level = level
        p.font.size = Pt(font_size - (level * 2))
        p.font.color.rgb = DARK
        p.font.name = "Calibri"
        p.space_after = Pt(10)


def add_two_column(slide, left_items, right_items, left_title=None, right_title=None):
    """Add two-column layout."""
    if left_title:
        lt = slide.shapes.add_textbox(Inches(0.6), Inches(1.5), Inches(4.2), Inches(0.4))
        lp = lt.text_frame.paragraphs[0]
        lp.text = left_title
        lp.font.size = Pt(22)
        lp.font.bold = True
        lp.font.color.rgb = PRIMARY
        lp.font.name = "Calibri"

    if right_title:
        rt = slide.shapes.add_textbox(Inches(5.2), Inches(1.5), Inches(4.2), Inches(0.4))
        rp = rt.text_frame.paragraphs[0]
        rp.text = right_title
        rp.font.size = Pt(22)
        rp.font.bold = True
        rp.font.color.rgb = SECONDARY
        rp.font.name = "Calibri"

    add_bullets(slide, left_items, left=0.6, top=2.0, width=4.2, height=4.5, font_size=18)
    add_bullets(slide, right_items, left=5.2, top=2.0, width=4.2, height=4.5, font_size=18)


def add_card(slide, left, top, width, height, title, body, title_color=PRIMARY):
    """Add a rounded card with title and body."""
    card = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height)
    )
    card.fill.solid()
    card.fill.fore_color.rgb = LIGHT_BG
    card.line.color.rgb = RGBColor(0xE5, 0xE7, 0xEB)

    tb = slide.shapes.add_textbox(
        Inches(left + 0.15), Inches(top + 0.15), Inches(width - 0.3), Inches(height - 0.3)
    )
    tf = tb.text_frame
    tf.word_wrap = True

    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = title_color
    p.font.name = "Calibri"

    p2 = tf.add_paragraph()
    p2.text = body
    p2.font.size = Pt(13)
    p2.font.color.rgb = DARK
    p2.font.name = "Calibri"
    p2.space_before = Pt(6)


def create_presentation():
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)

    # ── Slide 1: Title ──────────────────────────────────────────────
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank
    set_slide_bg(slide, PRIMARY)

    # Decorative accent bar
    accent = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0), Inches(3.2), Inches(0.15), Inches(1.5)
    )
    accent.fill.solid()
    accent.fill.fore_color.rgb = ACCENT
    accent.line.fill.background()

    title_box = slide.shapes.add_textbox(Inches(0.8), Inches(2.5), Inches(8.4), Inches(1.2))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.text = "Supervised Machine Learning"
    p.font.size = Pt(44)
    p.font.bold = True
    p.font.color.rgb = WHITE
    p.font.name = "Calibri"

    sub_box = slide.shapes.add_textbox(Inches(0.8), Inches(3.8), Inches(8.4), Inches(0.8))
    stf = sub_box.text_frame
    sp = stf.paragraphs[0]
    sp.text = "Learning from labeled data to make predictions"
    sp.font.size = Pt(22)
    sp.font.color.rgb = RGBColor(0xBF, 0xDB, 0xFE)
    sp.font.name = "Calibri"

    date_box = slide.shapes.add_textbox(Inches(0.8), Inches(6.2), Inches(4), Inches(0.4))
    dtf = date_box.text_frame
    dp = dtf.paragraphs[0]
    dp.text = "August 2026"
    dp.font.size = Pt(14)
    dp.font.color.rgb = RGBColor(0x93, 0xC5, 0xFD)
    dp.font.name = "Calibri"

    # ── Slide 2: What is Machine Learning? ────────────────────────
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, WHITE)
    add_title_bar(slide, "What is Machine Learning?")

    add_bullets(slide, [
        "Machine Learning (ML) enables computers to learn patterns from data — without being explicitly programmed for every rule",
        "Instead of writing \"if-then\" rules, we show the computer examples and let it figure out the pattern",
        "Three main types of ML:",
        ("Supervised Learning — learn from labeled examples", 1),
        ("Unsupervised Learning — find hidden patterns in unlabeled data", 1),
        ("Reinforcement Learning — learn by trial and error with rewards", 1),
    ], top=1.6, font_size=20)

    # Highlight box
    highlight = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6), Inches(5.8), Inches(8.8), Inches(0.9)
    )
    highlight.fill.solid()
    highlight.fill.fore_color.rgb = RGBColor(0xDB, 0xEA, 0xFE)
    highlight.line.fill.background()
    hb = slide.shapes.add_textbox(Inches(0.8), Inches(5.95), Inches(8.4), Inches(0.6))
    htf = hb.text_frame
    hp = htf.paragraphs[0]
    hp.text = "Today we focus on Supervised Learning — the most widely used approach in industry"
    hp.font.size = Pt(16)
    hp.font.bold = True
    hp.font.color.rgb = PRIMARY
    hp.font.name = "Calibri"

    # ── Slide 3: What is Supervised Learning? ─────────────────────
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, WHITE)
    add_title_bar(slide, "What is Supervised Learning?")

    add_bullets(slide, [
        "Supervised learning trains a model using input-output pairs (labeled data)",
        "The \"supervisor\" is the correct answer (label) provided for each training example",
        "Goal: learn a mapping function  f(x) → y  so the model can predict y for new, unseen x",
        "",
        "Analogy: Like studying for an exam with an answer key",
        ("You practice problems (training data) with solutions (labels)", 1),
        ("Then you solve new problems on your own (predictions)", 1),
    ], top=1.6, font_size=20)

    # ── Slide 4: Key Components ─────────────────────────────────────
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, WHITE)
    add_title_bar(slide, "Key Components")

    cards = [
        ("Features (X)", "Input variables used to make predictions.\nExample: house size, bedrooms, location", PRIMARY),
        ("Labels (Y)", "The correct output we want to predict.\nExample: house price, spam/not spam", SECONDARY),
        ("Training Data", "A dataset of (feature, label) pairs used to teach the model.", ACCENT),
        ("Model", "The learned function that maps features → predicted labels.", RGBColor(0x7C, 0x3A, 0xED)),
    ]
    for i, (title, body, color) in enumerate(cards):
        col = i % 2
        row = i // 2
        add_card(slide, 0.6 + col * 4.6, 1.6 + row * 2.5, 4.2, 2.2, title, body, color)

    # ── Slide 5: Two Types ──────────────────────────────────────────
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, WHITE)
    add_title_bar(slide, "Two Types of Supervised Learning")

    add_two_column(
        slide,
        [
            "Predicts continuous numerical values",
            "Output is a number on a scale",
            "Examples:",
            ("House prices ($250,000)", 1),
            ("Temperature tomorrow (72°F)", 1),
            ("Stock prices", 1),
            ("Sales revenue", 1),
        ],
        [
            "Predicts discrete categories or classes",
            "Output is a label from a fixed set",
            "Examples:",
            ("Spam vs. Not Spam", 1),
            ("Cat, Dog, or Bird", 1),
            ("Disease: Yes / No", 1),
            ("Sentiment: Positive / Negative", 1),
        ],
        left_title="Regression",
        right_title="Classification",
    )

    # ── Slide 6: Classification Deep Dive ─────────────────────────────
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, WHITE)
    add_title_bar(slide, "Classification in Detail", "Predicting categories")

    add_bullets(slide, [
        "Binary Classification — two classes (e.g., fraud / not fraud)",
        "Multi-class Classification — three or more classes (e.g., digit 0–9)",
        "Multi-label Classification — multiple labels per input (e.g., tags on a photo)",
        "",
        "How it works:",
        ("Model outputs probabilities for each class", 1),
        ("The class with the highest probability wins", 1),
        ("Decision boundary separates classes in feature space", 1),
    ], top=1.6, font_size=19)

    # ── Slide 7: Regression Deep Dive ───────────────────────────────
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, WHITE)
    add_title_bar(slide, "Regression in Detail", "Predicting continuous values")

    add_bullets(slide, [
        "Linear Regression — fits a straight line:  y = mx + b",
        "Polynomial Regression — fits curves for non-linear relationships",
        "Multiple Regression — uses several features simultaneously",
        "",
        "Example: Predicting house price",
        ("Features: sq ft, bedrooms, age, zip code", 1),
        ("Label: sale price in dollars", 1),
        ("Model learns weights for each feature", 1),
    ], top=1.6, font_size=19)

    # ── Slide 8: The Training Process ───────────────────────────────
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, WHITE)
    add_title_bar(slide, "The Training Process")

    steps = [
        ("1. Collect Data", "Gather labeled examples"),
        ("2. Preprocess", "Clean, normalize, split train/test"),
        ("3. Choose Model", "Pick an algorithm suited to the task"),
        ("4. Train", "Feed data, minimize prediction error"),
        ("5. Evaluate", "Test on held-out data"),
        ("6. Deploy", "Use model on real-world inputs"),
    ]
    for i, (title, body) in enumerate(steps):
        col = i % 3
        row = i // 3
        add_card(slide, 0.5 + col * 3.15, 1.5 + row * 2.6, 2.9, 2.2, title, body, PRIMARY)

    # ── Slide 9: Train / Validation / Test Split ────────────────────
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, WHITE)
    add_title_bar(slide, "Data Splitting Strategy")

    # Visual bars
    splits = [
        ("Training Set", "70%", PRIMARY, 0.6, 5.6),
        ("Validation Set", "15%", SECONDARY, 4.5, 1.5),
        ("Test Set", "15%", ACCENT, 6.3, 1.5),
    ]
    for name, pct, color, left, width in splits:
        bar = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(3.5), Inches(width), Inches(1.2)
        )
        bar.fill.solid()
        bar.fill.fore_color.rgb = color
        bar.line.fill.background()

        tb = slide.shapes.add_textbox(Inches(left), Inches(3.7), Inches(width), Inches(0.8))
        tf = tb.text_frame
        tf.paragraphs[0].alignment = PP_ALIGN.CENTER
        p = tf.paragraphs[0]
        p.text = f"{name}\n{pct}"
        p.font.size = Pt(14)
        p.font.bold = True
        p.font.color.rgb = WHITE
        p.font.name = "Calibri"

    add_bullets(slide, [
        "Training Set — used to learn model parameters",
        "Validation Set — tune hyperparameters, prevent overfitting",
        "Test Set — final unbiased evaluation (used only once!)",
        "Never train on test data — that leads to overly optimistic results",
    ], top=1.5, font_size=18)

    # ── Slide 10: Common Algorithms ───────────────────────────────────
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, WHITE)
    add_title_bar(slide, "Common Algorithms")

    algorithms = [
        ("Linear / Logistic Regression", "Simple, interpretable baseline", PRIMARY),
        ("Decision Trees", "Easy to visualize, handles non-linear data", SECONDARY),
        ("Random Forest", "Ensemble of trees, robust and accurate", ACCENT),
        ("Support Vector Machines", "Effective in high-dimensional spaces", RGBColor(0x7C, 0x3A, 0xED)),
        ("k-Nearest Neighbors", "Classify by similarity to neighbors", RGBColor(0xEC, 0x48, 0x99)),
        ("Neural Networks", "Deep learning for complex patterns", RGBColor(0xEF, 0x44, 0x44)),
    ]
    for i, (name, desc, color) in enumerate(algorithms):
        col = i % 2
        row = i // 2
        add_card(slide, 0.5 + col * 4.7, 1.5 + row * 1.85, 4.4, 1.6, name, desc, color)

    # ── Slide 11: Evaluation Metrics ────────────────────────────────
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, WHITE)
    add_title_bar(slide, "How Do We Measure Success?")

    add_two_column(
        slide,
        [
            "Mean Absolute Error (MAE)",
            "Mean Squared Error (MSE)",
            "Root Mean Squared Error (RMSE)",
            "R² Score (coefficient of determination)",
        ],
        [
            "Accuracy",
            "Precision & Recall",
            "F1 Score",
            "ROC-AUC Curve",
            "Confusion Matrix",
        ],
        left_title="Regression Metrics",
        right_title="Classification Metrics",
    )

    # ── Slide 12: Overfitting vs Underfitting ─────────────────────────
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, WHITE)
    add_title_bar(slide, "Overfitting vs. Underfitting")

    add_card(slide, 0.5, 1.6, 2.9, 2.5, "Underfitting", "Model is too simple\nFails to capture patterns\nHigh error on train AND test", SECONDARY)
    add_card(slide, 3.55, 1.6, 2.9, 2.5, "Good Fit", "Model captures true patterns\nLow error on both sets\nGeneralizes well", PRIMARY)
    add_card(slide, 6.6, 1.6, 2.9, 2.5, "Overfitting", "Model memorizes training data\nLow train error, high test error\nFails on new data", RGBColor(0xEF, 0x44, 0x44))

    add_bullets(slide, [
        "Solutions: more data, regularization, cross-validation, simpler models, early stopping",
    ], top=4.5, font_size=18)

    # ── Slide 13: Real-World Applications ───────────────────────────
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, WHITE)
    add_title_bar(slide, "Real-World Applications")

    apps = [
        ("Healthcare", "Disease diagnosis, drug discovery, patient risk scoring"),
        ("Finance", "Credit scoring, fraud detection, algorithmic trading"),
        ("Marketing", "Customer churn prediction, recommendation engines"),
        ("Technology", "Spam filters, voice assistants, image recognition"),
        ("Transportation", "Self-driving cars, route optimization, ETA prediction"),
        ("Education", "Personalized learning, automated grading"),
    ]
    for i, (domain, use) in enumerate(apps):
        col = i % 2
        row = i // 2
        add_card(slide, 0.5 + col * 4.7, 1.5 + row * 1.85, 4.4, 1.6, domain, use, PRIMARY)

    # ── Slide 14: Challenges ──────────────────────────────────────────
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, WHITE)
    add_title_bar(slide, "Challenges & Limitations")

    add_bullets(slide, [
        "Data Quality — garbage in, garbage out; labels must be accurate",
        "Bias & Fairness — models can perpetuate societal biases in training data",
        "Need for Labeled Data — labeling is expensive and time-consuming",
        "Interpretability — complex models (deep learning) are \"black boxes\"",
        "Distribution Shift — model performance drops when real-world data differs from training data",
        "Ethical Concerns — privacy, accountability, and responsible AI use",
    ], top=1.6, font_size=19)

    # ── Slide 15: Summary ───────────────────────────────────────────
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, PRIMARY)

    title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.8), Inches(8.4), Inches(0.8))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.text = "Key Takeaways"
    p.font.size = Pt(36)
    p.font.bold = True
    p.font.color.rgb = WHITE
    p.font.name = "Calibri"

    takeaways = [
        "Supervised ML learns from labeled (input, output) pairs",
        "Two main tasks: Classification (categories) and Regression (numbers)",
        "Process: collect data → train model → evaluate → deploy",
        "Choose the right algorithm and metrics for your problem",
        "Watch for overfitting and ensure data quality",
    ]
    for i, text in enumerate(takeaways):
        tb = slide.shapes.add_textbox(Inches(1.0), Inches(2.0 + i * 0.85), Inches(8.0), Inches(0.7))
        ttf = tb.text_frame
        tp = ttf.paragraphs[0]
        tp.text = f"✓  {text}"
        tp.font.size = Pt(20)
        tp.font.color.rgb = WHITE
        tp.font.name = "Calibri"

    thanks = slide.shapes.add_textbox(Inches(0.8), Inches(6.2), Inches(8.4), Inches(0.5))
    stf = thanks.text_frame
    sp = stf.paragraphs[0]
    sp.text = "Questions?"
    sp.font.size = Pt(28)
    sp.font.bold = True
    sp.font.color.rgb = ACCENT
    sp.font.name = "Calibri"
    sp.alignment = PP_ALIGN.CENTER

    return prs


if __name__ == "__main__":
    output_path = "/workspace/Supervised_Machine_Learning.pptx"
    prs = create_presentation()
    prs.save(output_path)
    print(f"Presentation saved to: {output_path}")
    print(f"Total slides: {len(prs.slides)}")
