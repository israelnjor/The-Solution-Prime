"""Create fixtures-preservation/ (E_preservation.docx, F_pictures.docx) with Microsoft Word.

Fixture E holds features Prime Word does not model yet — footnotes, comments, tracked insertions
and deletions, a content control, a table of contents, an equation, a bookmark with an internal
link, and a floating image. The engine must carry them through unchanged and warn about them.
"""
import os
from pathlib import Path

from generate_fixtures import chart_png
from word import Word

HERE = Path(__file__).resolve().parent
OUT = HERE / "fixtures-preservation"
WD_STYLE_NORMAL, WD_STYLE_HEADING_1, WD_STYLE_HEADING_2 = -1, -2, -3
WD_CONTENT_CONTROL_RICH_TEXT = 0


def build(app, path, image_path):
    doc = app.Documents.Add()

    def add(text, style=WD_STYLE_NORMAL):
        end = doc.Content.End - 1
        r = doc.Range(end, end)
        r.InsertAfter(text)
        r.Style = style
        out = doc.Range(r.Start, r.End)
        r.InsertParagraphAfter()
        return out

    add("Fixture E — Preservation [E-001]", WD_STYLE_HEADING_1)
    toc = add("Table of contents placeholder")
    add("Footnotes and comments [E-002]", WD_STYLE_HEADING_2)
    r = add("This sentence carries a footnote [E-003]")
    doc.Footnotes.Add(Range=doc.Range(r.End, r.End), Text="Footnote text [E-004]")
    r = add("This sentence carries a comment [E-005]")
    doc.Comments.Add(r, "Comment body [E-006]")

    add("Tracked changes [E-007]", WD_STYLE_HEADING_2)
    doc.TrackRevisions = True
    add("Inserted with change tracking on [E-008]")
    doc.TrackRevisions = False
    r = add("Deleted with change tracking on [E-009]")
    doc.TrackRevisions = True
    doc.Range(r.Start, r.End).Delete()
    doc.TrackRevisions = False

    add("Content control, bookmark and internal link [E-010]", WD_STYLE_HEADING_2)
    r = add("Text inside a rich-text content control [E-011]")
    doc.ContentControls.Add(WD_CONTENT_CONTROL_RICH_TEXT, r)
    r = add("Bookmarked paragraph [E-012]")
    doc.Bookmarks.Add("PrimeBookmark", r)
    r = add("Jump to the bookmarked paragraph [E-013]")
    doc.Hyperlinks.Add(r, "", "PrimeBookmark")

    add("Equation and floating image [E-014]", WD_STYLE_HEADING_2)
    r = add("a+b=c")
    doc.OMaths.Add(r)
    doc.OMaths(1).BuildUp()
    r = add("A floating image is anchored to this paragraph [E-015]")
    doc.Shapes.AddPicture(FileName=str(image_path), LinkToFile=False, SaveWithDocument=True,
                          Left=320, Top=40, Width=150, Height=75, Anchor=r)
    add("Closing paragraph [E-016]")

    doc.TablesOfContents.Add(Range=toc, UseHeadingStyles=True, UpperHeadingLevel=1, LowerHeadingLevel=2)
    doc.SaveAs2(FileName=str(path), FileFormat=16)
    doc.Close(0)


def build_pictures(app, path, image_path):
    """Fixture F — pictures: plain, bordered, cropped, alt text, in a table, and in the header."""
    doc = app.Documents.Add()

    def add(text, style=WD_STYLE_NORMAL):
        end = doc.Content.End - 1
        r = doc.Range(end, end)
        r.InsertAfter(text)
        r.Style = style
        out = doc.Range(r.Start, r.End)
        r.InsertParagraphAfter()
        return out

    def picture_after(r, width=150):
        at = doc.Range(r.End, r.End)
        shape = doc.InlineShapes.AddPicture(FileName=str(image_path), LinkToFile=False, SaveWithDocument=True, Range=at)
        shape.LockAspectRatio = True
        shape.Width = width
        return shape

    header = doc.Sections(1).Headers(1).Range
    header.Text = "Header with a logo [F-001] "
    end = doc.Sections(1).Headers(1).Range
    end.Collapse(0)  # wdCollapseEnd; Word refuses to resize header pictures via automation, so keep natural size
    doc.Sections(1).Headers(1).Range.InlineShapes.AddPicture(FileName=str(image_path), LinkToFile=False,
                                                             SaveWithDocument=True, Range=end)

    add("Fixture F — Pictures [F-002]", WD_STYLE_HEADING_1)
    picture_after(add("Plain inline picture: [F-003] "))
    bordered = picture_after(add("Picture with a red border: [F-004] "))
    bordered.Line.Visible = True
    bordered.Line.ForeColor.RGB = 0x0000C0  # BGR
    bordered.Line.Weight = 3
    cropped = picture_after(add("Cropped picture: [F-005] "))
    cropped.PictureFormat.CropLeft = 30
    cropped.PictureFormat.CropRight = 30
    described = picture_after(add("Picture with alt text: [F-006] "))
    described.AlternativeText = "Synthetic bar chart used as alt-text test"

    add("Picture inside a table [F-007]", WD_STYLE_HEADING_1)
    anchor = add("")
    table = doc.Tables.Add(anchor, 1, 2)
    table.Borders.Enable = True
    table.Cell(1, 1).Range.Text = "Logo cell [F-008]"
    table.Cell(1, 2).Range.InlineShapes.AddPicture(FileName=str(image_path), LinkToFile=False,
                                                    SaveWithDocument=True).Width = 80
    add("Closing paragraph [F-009]")
    doc.SaveAs2(FileName=str(path), FileFormat=16)
    doc.Close(0)


def main():
    OUT.mkdir(exist_ok=True)
    image = OUT / "_floating.png"
    image.write_bytes(chart_png(4, (300, 150)).getvalue())
    try:
        with Word() as w:
            build(w.app, OUT / "E_preservation.docx", image)
            build_pictures(w.app, OUT / "F_pictures.docx", image)
    finally:
        os.remove(image)
    print(f"wrote {OUT / 'E_preservation.docx'} and {OUT / 'F_pictures.docx'}")


if __name__ == "__main__":
    main()
