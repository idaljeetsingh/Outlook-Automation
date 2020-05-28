"""
    CONFIGURATION FOR AutomatingOutlook
"""

# Voting Options Semi-Colon[;] separated ex: "Yes;No"
VOTING_OPTIONS = "Yes;NO"

# Subject of the mail to automate further tasks
SUBJECT = "HELLO"

# Reply message if no attachment found for the matching subject
ERR_NO_ATTACHMENT_FOUND = """
                             Hello,
                             No attachments found or not in proper attachment not in proper extension for automating further tasks.
                          """

# Types of attachments allowed
ALLOWED_ATTACHMENT_TYPES = ['xls', 'xlsx']
