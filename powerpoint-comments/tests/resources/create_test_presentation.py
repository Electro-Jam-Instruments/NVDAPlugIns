"""
Create a test PowerPoint presentation about Guide Dogs
with comments and @mentions for testing the NVDA plugin.

Run this script with PowerPoint installed to generate the test deck.
"""

import win32com.client
import os

def create_guide_dog_presentation():
    """Create a test presentation about guide dogs with comments."""

    # Start PowerPoint
    ppt = win32com.client.Dispatch("PowerPoint.Application")
    ppt.Visible = True

    # Create new presentation
    presentation = ppt.Presentations.Add()

    # Get the output path
    script_dir = os.path.dirname(os.path.abspath(__file__))
    output_path = os.path.join(script_dir, "Guide_Dogs_Example_Deck.pptx")

    # ========== SLIDE 1: Title Slide ==========
    slide1 = presentation.Slides.Add(1, 1)  # ppLayoutTitle
    slide1.Shapes.Title.TextFrame.TextRange.Text = "Guide Dogs: Partners in Independence"
    slide1.Shapes(2).TextFrame.TextRange.Text = "A Comprehensive Overview\nTest Presentation for NVDA Plugin"

    # Add comment to title slide
    slide1.Comments.Add(
        Left=100,
        Top=100,
        Author="Sarah Johnson",
        AuthorInitials="SJ",
        Text="@John Smith please review the title - should we include the year?"
    )

    # ========== SLIDE 2: What Are Guide Dogs? ==========
    slide2 = presentation.Slides.Add(2, 2)  # ppLayoutText
    slide2.Shapes.Title.TextFrame.TextRange.Text = "What Are Guide Dogs?"

    bullet_text = """Guide dogs are specially trained assistance animals
They help people who are blind or visually impaired navigate safely
Training takes approximately 2 years from birth to placement
The most common breeds are Labrador Retrievers, Golden Retrievers, and German Shepherds
Guide dogs are protected under the Americans with Disabilities Act (ADA)"""

    slide2.Shapes(2).TextFrame.TextRange.Text = bullet_text

    # Add multiple comments
    slide2.Comments.Add(
        Left=150,
        Top=200,
        Author="John Smith",
        AuthorInitials="JS",
        Text="Great overview! Maybe add info about the handler relationship?"
    )

    slide2.Comments.Add(
        Left=200,
        Top=300,
        Author="Maria Garcia",
        AuthorInitials="MG",
        Text="@Sarah Johnson I think we should mention guide dog schools here too."
    )

    # ========== SLIDE 3: Guide Dog Statistics TABLE ==========
    slide3 = presentation.Slides.Add(3, 11)  # ppLayoutTitleOnly
    slide3.Shapes.Title.TextFrame.TextRange.Text = "Guide Dog Usage by Country (2024)"

    # Add a table
    countries_data = [
        ("Country", "Active Teams", "Primary Schools"),
        ("United States", "10,000", "Guide Dogs for the Blind, The Seeing Eye"),
        ("United Kingdom", "5,000", "Guide Dogs UK"),
        ("Germany", "3,500", "German Guide Dog Schools"),
        ("Australia", "2,800", "Guide Dogs Australia"),
        ("Canada", "2,500", "CNIB Guide Dogs"),
        ("France", "2,200", "Chiens Guides"),
        ("Japan", "1,800", "Japan Guide Dog Association"),
        ("Netherlands", "1,200", "KNGF Geleidehonden"),
    ]

    num_rows = len(countries_data)
    num_cols = 3
    table_left = 50
    table_top = 120
    table_width = 620
    table_height = 320

    table_shape = slide3.Shapes.AddTable(num_rows, num_cols, table_left, table_top, table_width, table_height)
    table = table_shape.Table

    # Populate the table
    for row_idx, row_data in enumerate(countries_data):
        for col_idx, cell_value in enumerate(row_data):
            table.Cell(row_idx + 1, col_idx + 1).Shape.TextFrame.TextRange.Text = cell_value

    # Add comments about the table
    slide3.Comments.Add(
        Left=550,
        Top=150,
        Author="David Chen",
        AuthorInitials="DC",
        Text="These numbers are estimates. @Maria Garcia can you verify the UK data?"
    )

    slide3.Comments.Add(
        Left=550,
        Top=250,
        Author="Sarah Johnson",
        AuthorInitials="SJ",
        Text="Should we add a note about data sources?"
    )

    # ========== SLIDE 4: Guide Dog Breeds (Visual representation) ==========
    slide4 = presentation.Slides.Add(4, 2)  # ppLayoutText
    slide4.Shapes.Title.TextFrame.TextRange.Text = "Most Common Guide Dog Breeds"

    # Text-based chart representation (more reliable than COM chart)
    chart_text = """BREED DISTRIBUTION (Approximate %)

Labrador Retriever:  ████████████████████████████████████  70%
Golden Retriever:    ████████                              15%
German Shepherd:     █████                                 10%
Other Breeds:        ███                                    5%

Total Active Guide Dogs Worldwide: ~30,000 teams"""

    slide4.Shapes(2).TextFrame.TextRange.Text = chart_text
    slide4.Shapes(2).TextFrame.TextRange.Font.Name = "Consolas"
    slide4.Shapes(2).TextFrame.TextRange.Font.Size = 14

    # Add comment about the chart
    slide4.Comments.Add(
        Left=550,
        Top=200,
        Author="Maria Garcia",
        AuthorInitials="MG",
        Text="@John Smith can you double-check these percentages? I got them from the 2023 IGDF report."
    )

    # ========== SLIDE 5: Training Process ==========
    slide5 = presentation.Slides.Add(5, 2)  # ppLayoutText
    slide5.Shapes.Title.TextFrame.TextRange.Text = "The Guide Dog Training Process"

    training_text = """Puppy Raising (0-18 months): Volunteer families socialize puppies
Formal Training (4-6 months): Professional trainers teach navigation skills
Matching: Dogs are paired with handlers based on lifestyle and walking speed
Team Training (2-4 weeks): Handler and dog learn to work together
Follow-up Support: Ongoing assistance from the guide dog school"""

    slide5.Shapes(2).TextFrame.TextRange.Text = training_text

    # Comment with reply simulation (threaded comments)
    slide5.Comments.Add(
        Left=100,
        Top=150,
        Author="John Smith",
        AuthorInitials="JS",
        Text="@David Chen what's the success rate for formal training?"
    )

    # ========== SLIDE 6: Benefits ==========
    slide6 = presentation.Slides.Add(6, 2)  # ppLayoutText
    slide6.Shapes.Title.TextFrame.TextRange.Text = "Benefits of Guide Dogs"

    benefits_text = """Increased mobility and independence
Enhanced safety when navigating obstacles and traffic
Improved confidence in unfamiliar environments
Companionship and emotional support
Greater participation in work, education, and social activities
Reduced reliance on sighted guides"""

    slide6.Shapes(2).TextFrame.TextRange.Text = benefits_text

    # No comments on this slide - good for testing "no comments" scenario

    # ========== SLIDE 7: How to Interact ==========
    slide7 = presentation.Slides.Add(7, 2)  # ppLayoutText
    slide7.Shapes.Title.TextFrame.TextRange.Text = "Etiquette: Interacting with Guide Dog Teams"

    etiquette_text = """DO NOT pet, feed, or distract a working guide dog
Always speak to the handler, not the dog
Ask before offering assistance
Give the team plenty of space to navigate
Never grab the harness or the handler's arm
It's okay to ask questions politely!"""

    slide7.Shapes(2).TextFrame.TextRange.Text = etiquette_text

    # Multiple comments including @mentions
    slide7.Comments.Add(
        Left=100,
        Top=200,
        Author="Maria Garcia",
        AuthorInitials="MG",
        Text="This is really important information for the general public."
    )

    slide7.Comments.Add(
        Left=150,
        Top=280,
        Author="Sarah Johnson",
        AuthorInitials="SJ",
        Text="@John Smith @Maria Garcia should we add a section about access rights too?"
    )

    slide7.Comments.Add(
        Left=200,
        Top=350,
        Author="David Chen",
        AuthorInitials="DC",
        Text="Great suggestion @Sarah Johnson - maybe on a separate slide?"
    )

    # ========== SLIDE 8: Resources ==========
    slide8 = presentation.Slides.Add(8, 2)  # ppLayoutText
    slide8.Shapes.Title.TextFrame.TextRange.Text = "Guide Dog Organizations"

    resources_text = """Guide Dogs for the Blind (USA)
The Seeing Eye (USA) - First guide dog school in America
Guide Dogs UK
Seeing Eye Dogs Australia
Canadian Guide Dogs for the Blind
International Guide Dog Federation - 90+ member organizations worldwide"""

    slide8.Shapes(2).TextFrame.TextRange.Text = resources_text

    slide8.Comments.Add(
        Left=100,
        Top=150,
        Author="John Smith",
        AuthorInitials="JS",
        Text="We should add links to these organizations in the notes."
    )

    # ========== SLIDE 9: Thank You ==========
    slide9 = presentation.Slides.Add(9, 1)  # ppLayoutTitle
    slide9.Shapes.Title.TextFrame.TextRange.Text = "Thank You!"
    slide9.Shapes(2).TextFrame.TextRange.Text = "Questions?\n\nContact: guidedogs@example.com"

    slide9.Comments.Add(
        Left=100,
        Top=100,
        Author="Sarah Johnson",
        AuthorInitials="SJ",
        Text="@John Smith @Maria Garcia @David Chen - Final review complete?"
    )

    # Save the presentation
    presentation.SaveAs(output_path)

    print(f"\n{'='*60}")
    print("TEST PRESENTATION CREATED SUCCESSFULLY!")
    print(f"{'='*60}")
    print(f"\nSaved to: {output_path}")
    print(f"\nPresentation contains:")
    print(f"  - 9 slides about guide dogs")
    print(f"  - 1 TABLE (Slide 3: Country statistics)")
    print(f"  - 1 Text-based chart (Slide 4: Breed distribution)")
    print(f"  - 12 comments total across slides")
    print(f"  - Multiple @mentions (Sarah Johnson, John Smith, Maria Garcia, David Chen)")
    print(f"\nSlide comment breakdown:")
    print(f"  Slide 1: Title - 1 comment (with @mention)")
    print(f"  Slide 2: What Are Guide Dogs - 2 comments (1 with @mention)")
    print(f"  Slide 3: TABLE - 2 comments (1 with @mention)")
    print(f"  Slide 4: CHART - 1 comment (with @mention)")
    print(f"  Slide 5: Training Process - 1 comment (with @mention)")
    print(f"  Slide 6: Benefits - 0 comments (test empty case)")
    print(f"  Slide 7: Etiquette - 3 comments (multiple @mentions)")
    print(f"  Slide 8: Resources - 1 comment")
    print(f"  Slide 9: Thank You - 1 comment (multiple @mentions)")
    print(f"\n{'='*60}")

    return output_path


if __name__ == "__main__":
    try:
        path = create_guide_dog_presentation()
        print("\nPresentation is now open in PowerPoint.")
        print("You can use it to test the NVDA plugin!")
    except Exception as e:
        print(f"Error creating presentation: {e}")
        import traceback
        traceback.print_exc()
