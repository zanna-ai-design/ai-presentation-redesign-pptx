from PPTX_style import (
    create_presentation, add_blank_slide,
    add_bg, add_header_footer, add_rect,
    add_image, add_textbox, add_p, add_slide_title,
    SW, SH, PX, W,
    HEADER_H, FOOTER_H, TEXT_Y_OFFSET,
    MARGIN_X, CONTENT_Y, CONTENT_H,
    COL1_X, COL1_W, COL2_X, COL2_W, COL_GAP,
    FONT_TITLE, FONT_SUBHEAD, FONT_BODY,
    C_ACCENT, C_BODY, C_LIGHT, C_WHITE
)

prs = create_presentation()

# ─────────────────────────────────────
# SLIDE 1 — COVER
# ─────────────────────────────────────
slide1 = add_blank_slide(prs)
add_bg(slide1, color=C_ACCENT)

add_image(slide1, "piro4d-lightbulb-2632075_1280.jpg", 200, 162, 760, 473)
add_image(slide1, "GG_cover_header.png",                1020, 162, 710, 473)

add_rect(slide1, 0, SH - 162, SW, 162, C_LIGHT)
add_image(slide1, "GG_logo_cover.png",          200,  938, 414, 42)
add_image(slide1, "GG_slogan_white_cover.png",  1020, 938, 709, 50)

# ─────────────────────────────────────
# SLIDE 2 — CHART
# ─────────────────────────────────────
slide2 = add_blank_slide(prs)
add_bg(slide2)
add_header_footer(slide2)

add_slide_title(slide2, "DATA VISUALIZATION: FROM NOISE TO CLARITY")

text_y = CONTENT_Y + TEXT_Y_OFFSET

# text left
tf2 = add_textbox(slide2, COL1_X, text_y, COL1_W, CONTENT_H - 70)
#add_p(tf2, "", size=FONT_SUBHEAD)  # двойной отступ между блоками
add_p(tf2, "THE PROBLEM", bold=True, size=FONT_SUBHEAD, color=C_ACCENT, first=True)
add_p(tf2, "AI-generated charts use arbitrary high-saturation colors, retain default axis styling, and are inserted without visual integration into the layout.")
add_p(tf2, "")
add_p(tf2, "THE SOLUTION", bold=True, size=FONT_SUBHEAD, color=C_ACCENT)
add_p(tf2, "A refined corporate palette, clean axes, and proportional sizing make the chart a natural part of the visual system rather than a standalone element.")

# chart right
chart_w = COL2_W
chart_h = round(chart_w * (473 / 760))
add_image(slide2, "chart_after.png", COL2_X, text_y, chart_w, chart_h)

# ─────────────────────────────────────
# SLIDE 3 — TEXT
# ─────────────────────────────────────
slide3 = add_blank_slide(prs)
add_bg(slide3)
add_header_footer(slide3)

add_slide_title(slide3, "FROM AI OUTPUT TO EDITORIAL QUALITY")

# равные колонки
COL_W = (SW - MARGIN_X * 2 - COL_GAP) // 2
col1_x = MARGIN_X
col2_x = MARGIN_X + COL_W + COL_GAP

text_y3 = CONTENT_Y + TEXT_Y_OFFSET

# column 1
tf3a = add_textbox(slide3, col1_x, text_y3, COL_W, CONTENT_H - TEXT_Y_OFFSET)
add_p(tf3a, "TYPOGRAPHY AND LAYOUT", bold=True, size=FONT_SUBHEAD, color=C_ACCENT, first=True)
add_p(tf3a, "Default fonts and uniform sizing were replaced with a structured typographic hierarchy using Inter. Heading weight, size, and spacing now guide the reader clearly through the content. The result is a layout that feels intentional and easy to navigate.")
add_p(tf3a, "")
add_p(tf3a, "DATA PRESENTATION", bold=True, size=FONT_SUBHEAD, color=C_ACCENT)
add_p(tf3a, "The chart was redesigned with a refined color palette, clean axes, and integrated label styling. It now functions as part of the visual system rather than a standalone insert. Data becomes easier to read and visually consistent with the rest of the slide.")

# column 2
tf3b = add_textbox(slide3, col2_x, text_y3, COL_W, CONTENT_H - TEXT_Y_OFFSET)
add_p(tf3b, "COLOR AND VISUAL STYLE", bold=True, size=FONT_SUBHEAD, color=C_ACCENT, first=True)
add_p(tf3b, "Arbitrary default colors were replaced with a consistent corporate palette. All visual elements follow the same color logic, creating coherence across the presentation. Color now carries meaning rather than noise.")
add_p(tf3b, "")
add_p(tf3b, "BRAND IDENTITY", bold=True, size=FONT_SUBHEAD, color=C_ACCENT)
add_p(tf3b, "Every slide now carries the GreenGrid visual identity — logo, corporate palette, typography system, and consistent layout grid. The result is a document that represents the organization behind it, communicates with authority, and leaves a lasting professional impression.")

# ─────────────────────────────────────
# SLIDE 4 — CLOSING
# ─────────────────────────────────────
slide4 = add_blank_slide(prs)
add_bg(slide4, color=C_ACCENT)

# logo — bigger, vertically centered
add_image(slide4, "GG_logo_inversed.png", MARGIN_X, 420, 600, 62)

# address — bigger, more space from logo
tf4 = add_textbox(slide4, MARGIN_X, 510, 900, 80)
add_p(tf4, "47 Lumeris Avenue, Solarwind District,\nArvandor Republic",
      size=14, color=C_LIGHT, first=True)

# footer
add_rect(slide4, 0, SH - FOOTER_H, SW, FOOTER_H, C_LIGHT)
add_image(slide4, "GG_slogan_white_cover.png",
          SW - MARGIN_X - 709, SH - FOOTER_H + 18, 709, 50)

# ─────────────────────────────────────
# SAVE
# ─────────────────────────────────────
prs.save("presentation_after.pptx")