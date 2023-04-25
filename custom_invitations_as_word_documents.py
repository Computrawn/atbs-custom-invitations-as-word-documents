#! python3
# custom_invitations_as_word_documents.py — An exercise in manipulating Word documents.
# For more information, see project_details.txt.

import logging

logging.basicConfig(
    level=logging.DEBUG,
    filename="logging.txt",
    format="%(asctime)s -  %(levelname)s -  %(message)s",
)
logging.disable(logging.CRITICAL)  # Note out to enable logging.
