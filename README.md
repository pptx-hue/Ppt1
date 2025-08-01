from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.dml.color import RGBColor
from pathlib import Path

# Create a new presentation
prs = Presentation()
title_slide_layout = prs.slide_layouts[0]
content_slide_layout = prs.slide_layouts[1]

# Helper function to add a slide with title and content
def add_slide(title, content):
    slide = prs.slides.add_slide(content_slide_layout)
    slide.shapes.title.text = title
    content_shape = slide.placeholders[1]
    content_shape.text = content

# Slide 1: Student Intro
slide = prs.slides.add_slide(content_slide_layout)
slide.shapes.title.text = "Student Information"
student_info = """\
1. NAME: MD SUFIAN AKHTER
2. Roll No: 12405923043
3. REG NO: 231240210043
4. YEAR: 3rd YEAR
5. SEMESTER: 5th SEMESTER
6. SUBJECT: PHARMACOLOGY-II
7. PAPER CODE: PT518
8. COLLEGE NAME: GUPTA COLLEGE OF TECHNOLOGICAL SCIENCES"""
slide.placeholders[1].text = student_info

# Slide 2: Introduction to Histamine
add_slide("Introduction to Histamine", 
          "Histamine is a biogenic amine involved in local immune responses, regulating physiological function in the gut, and acting as a neurotransmitter. It is derived from the amino acid histidine.")

# Slide 3: Biosynthesis and Storage
add_slide("Biosynthesis and Storage", 
          "Histamine is synthesized by decarboxylation of histidine via the enzyme histidine decarboxylase. It is stored mainly in mast cells and basophils in granules and released during allergic reactions.")

# Slide 4: Mechanism of Action
add_slide("Mechanism of Action", 
          "Histamine acts by binding to histamine receptors (H1, H2, H3, H4) present on various cells. The binding triggers different cellular responses such as vasodilation, bronchoconstriction, and gastric acid secretion.")

# Slide 5: Physiological Roles
add_slide("Physiological Roles", 
          "• Vasodilation and increased vascular permeability\n• Smooth muscle contraction (e.g., in bronchi)\n• Stimulation of gastric acid secretion\n• Modulation of neurotransmission in the brain")

# Slide 6: Histamine Receptors (H1–H4)
add_slide("Histamine Receptors", 
          "• H1: Allergic responses, bronchoconstriction, vasodilation\n• H2: Gastric acid secretion\n• H3: Neurotransmitter modulation\n• H4: Immune cell chemotaxis and inflammation")

# Slide 7: Pathological Roles of Histamine
add_slide("Pathological Roles", 
          "• Allergic rhinitis\n• Urticaria (hives)\n• Asthma\n• Anaphylaxis\n• Gastric ulcers (due to excessive HCl secretion)")

# Slide 8: Antihistamines Overview
add_slide("Antihistamines", 
          "• H1 blockers: Used for allergies (e.g., diphenhydramine, loratadine)\n• H2 blockers: Reduce gastric acid (e.g., ranitidine, famotidine)\n• H3 & H4 antagonists: Under research for neurological and inflammatory conditions")

# Slide 9: Therapeutic Uses & Side Effects
add_slide("Therapeutic Uses & Side Effects", 
          "• Uses: Allergy relief, acid reflux treatment, motion sickness\n• Side Effects: Drowsiness, dry mouth, dizziness, GI disturbances")

# Slide 10: Conclusion
add_slide("Conclusion", 
          "Histamine plays a critical role in immune response, gastric secretion, and neurotransmission. Understanding its functions and receptor types helps in the effective use of antihistamines in clinical settings.")

# Save as PDF
output_path = "/mnt/data/Histamine_Presentation.pdf"
prs.save(output_path)

output_path
