"""
Example Python client for the DOCX Anonymization API

This script demonstrates how to interact with the API programmatically.
"""

import requests
import os
import sys
from pathlib import Path


class DocxAnonymizerClient:
    """
    Simple client for the DOCX Anonymization API.
    """
    
    def __init__(self, base_url="http://localhost:8000"):
        """
        Initialize the client.
        
        Args:
            base_url: Base URL of the API server
        """
        self.base_url = base_url.rstrip('/')
        self.anonymize_endpoint = f"{self.base_url}/anonymise-docx"
        self.health_endpoint = f"{self.base_url}/health"
    
    def health_check(self):
        """
        Check if the API is healthy and available.
        
        Returns:
            True if API is healthy, False otherwise
        """
        try:
            response = requests.get(self.health_endpoint, timeout=5)
            if response.status_code == 200:
                data = response.json()
                print(f"✓ API is healthy: {data}")
                return True
            else:
                print(f"✗ API health check failed: {response.status_code}")
                return False
        except requests.exceptions.RequestException as e:
            print(f"✗ API is not reachable: {e}")
            return False
    
    def anonymize_file(self, input_path, output_path=None):
        """
        Anonymize a DOCX file.
        
        Args:
            input_path: Path to input DOCX file
            output_path: Path to save anonymized file (optional)
            
        Returns:
            Path to anonymized file if successful, None otherwise
        """
        # Validate input file
        input_file = Path(input_path)
        if not input_file.exists():
            print(f"✗ Input file does not exist: {input_path}")
            return None
        
        if not input_file.suffix.lower() == '.docx':
            print(f"✗ Input file must be a .docx file: {input_path}")
            return None
        
        # Generate output path if not provided
        if output_path is None:
            output_path = input_file.parent / f"{input_file.stem}_anonymized.docx"
        
        # Prepare file upload
        print(f"📤 Uploading: {input_path}")
        
        with open(input_path, 'rb') as f:
            files = {
                'file': (input_file.name, f, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
            }
            
            try:
                # Send request
                response = requests.post(
                    self.anonymize_endpoint,
                    files=files,
                    timeout=30  # 30 second timeout
                )
                
                # Check response
                if response.status_code == 200:
                    # Save anonymized file
                    with open(output_path, 'wb') as out_f:
                        out_f.write(response.content)
                    
                    print(f"✓ Anonymized file saved to: {output_path}")
                    return str(output_path)
                
                else:
                    print(f"✗ API returned error {response.status_code}: {response.text}")
                    return None
                    
            except requests.exceptions.Timeout:
                print("✗ Request timed out. The file may be too large or the server is busy.")
                return None
                
            except requests.exceptions.RequestException as e:
                print(f"✗ Request failed: {e}")
                return None
    
    def batch_anonymize(self, input_dir, output_dir=None):
        """
        Anonymize all DOCX files in a directory.
        
        Args:
            input_dir: Directory containing DOCX files
            output_dir: Directory to save anonymized files (optional)
            
        Returns:
            List of successfully anonymized file paths
        """
        input_path = Path(input_dir)
        
        if not input_path.is_dir():
            print(f"✗ Input directory does not exist: {input_dir}")
            return []
        
        # Get all DOCX files
        docx_files = list(input_path.glob("*.docx"))
        
        if not docx_files:
            print(f"✗ No .docx files found in: {input_dir}")
            return []
        
        print(f"\n📁 Found {len(docx_files)} DOCX files")
        print("="*60)
        
        # Set output directory
        if output_dir is None:
            output_path = input_path / "anonymized"
        else:
            output_path = Path(output_dir)
        
        # Create output directory
        output_path.mkdir(exist_ok=True)
        
        # Process each file
        successful = []
        for i, docx_file in enumerate(docx_files, 1):
            print(f"\n[{i}/{len(docx_files)}] Processing: {docx_file.name}")
            
            output_file = output_path / f"{docx_file.stem}_anonymized.docx"
            result = self.anonymize_file(docx_file, output_file)
            
            if result:
                successful.append(result)
        
        print("\n" + "="*60)
        print(f"✓ Successfully anonymized {len(successful)}/{len(docx_files)} files")
        print(f"📁 Output directory: {output_path}")
        
        return successful


def main():
    """
    Main function with CLI interface.
    """
    print("\n" + "="*60)
    print("DOCX Anonymization API - Python Client")
    print("="*60 + "\n")
    
    # Check command line arguments
    if len(sys.argv) < 2:
        print("Usage:")
        print("  python api_client.py <input_file.docx> [output_file.docx]")
        print("  python api_client.py --batch <input_directory> [output_directory]")
        print("\nExamples:")
        print("  python api_client.py paper.docx")
        print("  python api_client.py paper.docx paper_anon.docx")
        print("  python api_client.py --batch ./papers ./anonymized")
        sys.exit(1)
    
    # Initialize client
    client = DocxAnonymizerClient()
    
    # Health check
    print("Checking API health...")
    if not client.health_check():
        print("\n✗ API is not available. Make sure the server is running:")
        print("  uvicorn app:app --host 0.0.0.0 --port 8000")
        sys.exit(1)
    
    print()
    
    # Batch mode
    if sys.argv[1] == "--batch":
        if len(sys.argv) < 3:
            print("✗ Please provide input directory for batch mode")
            sys.exit(1)
        
        input_dir = sys.argv[2]
        output_dir = sys.argv[3] if len(sys.argv) > 3 else None
        
        client.batch_anonymize(input_dir, output_dir)
    
    # Single file mode
    else:
        input_file = sys.argv[1]
        output_file = sys.argv[2] if len(sys.argv) > 2 else None
        
        result = client.anonymize_file(input_file, output_file)
        
        if result:
            print("\n✨ Success!")
        else:
            print("\n✗ Failed to anonymize file")
            sys.exit(1)


if __name__ == "__main__":
    main()
