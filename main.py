"""
Safe2Day - Main Application Entry Point
Orchestrates the logo selection and document formatting process.
IMPORTANT: Source documents in Forms folder are NEVER modified.
"""

import sys
from pathlib import Path

# Import our modules
import select_logo
import format_documents


def main():
    """
    Main workflow:
    1. Run select_logo.py to choose a client logo
    2. Run format_documents.py to create client copies with logo and PDFs
    
    GUARANTEE: Source documents in Forms folder are NEVER modified.
    All edits are done on copies in Client_Forms folder.
    """
    print("=" * 60)
    print("Safe2Day Document Processing System")
    print("=" * 60)
    print()
    print("This tool will:")
    print("  1. Let you select a client logo")
    print("  2. Copy S2D forms to Client_Forms folder")
    print("  3. Replace logos in the copies (NOT source documents)")
    print("  4. Create PDF versions of all client documents")
    print()
    print("⚠ IMPORTANT: Source documents in Forms folder are NEVER modified")
    print()
    print("=" * 60)
    print()
    
    # Step 1: Select logo
    print("STEP 1: Logo Selection")
    print("-" * 60)
    selected_logo = select_logo.select_logo()
    
    if not selected_logo:
        print("\n✗ No logo selected. Exiting.")
        sys.exit(1)
    
    print()
    
    # Step 2: Format documents and create PDFs
    print("STEP 2: Document Formatting")
    print("-" * 60)
    format_documents.process_documents()
    
    print()
    print("=" * 60)
    print("✓ Process Complete!")
    print("=" * 60)
    print()
    print("Next steps:")
    print("  - Check Client_Forms folder for Word documents with new logo")
    print("  - Check PDFs folder for PDF versions")
    print("  - Source documents in Forms folder remain unchanged")
    print()


if __name__ == "__main__":
    main()

