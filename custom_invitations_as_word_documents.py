#! python3
# custom_invitations_as_word_documents.py — An exercise in manipulating Word documents.
# For more information, see project_details.txt.

import logging
import docx

logging.basicConfig(
    level=logging.DEBUG,
    filename="logging.txt",
    format="%(asctime)s -  %(levelname)s -  %(message)s",
)
# logging.disable(logging.CRITICAL)  # Note out to enable logging.


def get_guestlist():
    """Read contents of 'guests.txt' and return list of guests."""
    with open("guests.txt", "r", encoding="utf-8") as f:
        guests = f.readlines()
        stripped_list = [guest.strip() for guest in guests]
    return stripped_list


def make_invitations(guests):
    """Open styled blank docx, write invitation for
    each guest on separate page and save file."""
    doc = docx.Document("template.docx")

    for idx, guest in enumerate(guests):
        doc.add_paragraph(
            "It would be a pleasure to have the company of", style="lines135"
        )
        doc.add_paragraph(guest, style="guest")
        doc.add_paragraph("at 11010 Memory Lane on the Evening of", style="lines135")
        doc.add_paragraph("April 1st", style="date")
        doc.add_paragraph("at 7 o'clock", style="lines135")
        if idx != len(guests) - 1:
            doc.add_page_break()

    doc.save("invitations.docx")


guest_list = get_guestlist()
make_invitations(guest_list)
