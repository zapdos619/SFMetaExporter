"""
ContentDocument export functionality - exports metadata and downloads files
"""
import os
import csv
import requests
from typing import List, Dict, Tuple
from salesforce_client import SalesforceClient


class ContentDocumentExporter:
    """Handles ContentDocument metadata export and file downloads from Salesforce"""
    
    def __init__(self, sf_client: SalesforceClient):
        """Initialize with Salesforce client"""
        self.sf_client = sf_client
        self.sf = sf_client.sf
        self.base_url = sf_client.base_url
        self.headers = sf_client.headers
    
    def export_content_documents(self, output_path: str) -> Tuple[str, Dict]:
        """
        Export ContentDocument metadata to CSV and download all files
        
        Args:
            output_path: Path for the CSV file
            
        Returns:
            Tuple of (csv_path, statistics_dict)
        """
        self._log_status("=== Starting ContentDocument Export ===")
        
        stats = {
            'total_documents': 0,
            'successful_downloads': 0,
            'failed_downloads': 0,
            'total_size_bytes': 0,
            'failed_files': []
        }
        
        # Create Documents folder in same directory as CSV
        csv_dir = os.path.dirname(output_path)
        documents_folder = os.path.join(csv_dir, "Documents")
        
        if not os.path.exists(documents_folder):
            os.makedirs(documents_folder)
            self._log_status(f"Created folder: {documents_folder}")
        
        # Query all ContentDocuments
        self._log_status("Querying ContentDocument records...")
        content_documents = self._query_content_documents()
        stats['total_documents'] = len(content_documents)
        
        self._log_status(f"Found {len(content_documents)} ContentDocument records")
        
        if len(content_documents) == 0:
            self._log_status("No ContentDocument records found in org")
            # Still create empty CSV
            self._create_csv_file([], output_path)
            return output_path, stats
        
        # Download each file
        for i, doc in enumerate(content_documents, 1):
            doc_id = doc['Id']
            title = doc['Title']
            file_extension = doc.get('FileExtension', '')
            file_type = doc.get('FileType', '')
            content_size = doc.get('ContentSize', 0)
            
            # Construct filename with extension
            if file_extension:
                filename = f"{title}.{file_extension}"
            else:
                filename = title
            
            self._log_status(f"[{i}/{len(content_documents)}] Downloading: {filename}")
            
            try:
                # Download the file
                file_path = self._download_file(doc_id, filename, documents_folder)
                stats['successful_downloads'] += 1
                stats['total_size_bytes'] += content_size
                self._log_status(f"  ✅ Downloaded to: {file_path}")
            except Exception as e:
                error_msg = str(e)
                self._log_status(f"  ❌ ERROR: {error_msg}")
                stats['failed_downloads'] += 1
                stats['failed_files'].append({
                    'filename': filename,
                    'id': doc_id,
                    'reason': error_msg
                })
        
        # Create CSV with metadata
        self._log_status("\n=== Creating CSV File ===")
        final_output_path = self._create_csv_file(content_documents, output_path)
        
        return final_output_path, stats
    
    def _query_content_documents(self) -> List[Dict]:
        """Query all ContentDocument records with standard fields"""
        try:
            # Query ContentDocument with standard fields
            query = """
                SELECT Id, Title, FileExtension, FileType, ContentSize, 
                       CreatedDate, CreatedById, LastModifiedDate, LastModifiedById,
                       OwnerId, ParentId, IsArchived, IsDeleted, 
                       ArchivedDate, ArchivedById, Description,
                       PublishStatus, LatestPublishedVersionId
                FROM ContentDocument
                ORDER BY CreatedDate DESC
            """
            
            result = self.sf.query_all(query)
            return result['records']
            
        except Exception as e:
            self._log_status(f"ERROR querying ContentDocument: {str(e)}")
            raise
    
    def _download_file(self, document_id: str, filename: str, destination_folder: str) -> str:
        """
        Download a single file from Salesforce
        
        Args:
            document_id: ContentDocument Id
            filename: Name for the downloaded file
            destination_folder: Folder to save the file
            
        Returns:
            Full path of downloaded file
        """
        try:
            # First, get the latest ContentVersion for this document
            version_query = f"SELECT Id, VersionData FROM ContentVersion WHERE ContentDocumentId = '{document_id}' AND IsLatest = true"
            version_result = self.sf.query(version_query)
            
            if not version_result['records']:
                raise Exception("No ContentVersion found")
            
            version_id = version_result['records'][0]['Id']
            
            # Download the file using VersionData
            download_url = f"{self.base_url}/services/data/v{self.sf_client.api_version}/sobjects/ContentVersion/{version_id}/VersionData"
            
            response = requests.get(download_url, headers=self.headers, timeout=120)
            response.raise_for_status()
            
            # Sanitize filename to remove invalid characters
            safe_filename = self._sanitize_filename(filename)
            file_path = os.path.join(destination_folder, safe_filename)
            
            # Handle duplicate filenames
            file_path = self._get_unique_filepath(file_path)
            
            # Write file to disk
            with open(file_path, 'wb') as f:
                f.write(response.content)
            
            return file_path
            
        except Exception as e:
            raise Exception(f"Failed to download: {str(e)}")
    
    def _sanitize_filename(self, filename: str) -> str:
        """Remove invalid characters from filename"""
        invalid_chars = '<>:"/\\|?*'
        for char in invalid_chars:
            filename = filename.replace(char, '_')
        return filename
    
    def _get_unique_filepath(self, filepath: str) -> str:
        """Generate unique filepath if file already exists"""
        if not os.path.exists(filepath):
            return filepath
        
        base, extension = os.path.splitext(filepath)
        counter = 1
        
        while os.path.exists(f"{base}_{counter}{extension}"):
            counter += 1
        
        return f"{base}_{counter}{extension}"
    
    def _create_csv_file(self, content_documents: List[Dict], output_path: str) -> str:
        """Create CSV file with ContentDocument metadata"""
        headers = [
            'Id', 'Title', 'FileExtension', 'FileType', 'ContentSize (Bytes)',
            'CreatedDate', 'CreatedById', 'LastModifiedDate', 'LastModifiedById',
            'OwnerId', 'ParentId', 'IsArchived', 'IsDeleted',
            'ArchivedDate', 'ArchivedById', 'Description',
            'PublishStatus', 'LatestPublishedVersionId'
        ]
        
        with open(output_path, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(headers)
            
            for doc in content_documents:
                row = [
                    doc.get('Id', ''),
                    doc.get('Title', ''),
                    doc.get('FileExtension', ''),
                    doc.get('FileType', ''),
                    doc.get('ContentSize', 0),
                    doc.get('CreatedDate', ''),
                    doc.get('CreatedById', ''),
                    doc.get('LastModifiedDate', ''),
                    doc.get('LastModifiedById', ''),
                    doc.get('OwnerId', ''),
                    doc.get('ParentId', ''),
                    doc.get('IsArchived', False),
                    doc.get('IsDeleted', False),
                    doc.get('ArchivedDate', ''),
                    doc.get('ArchivedById', ''),
                    doc.get('Description', ''),
                    doc.get('PublishStatus', ''),
                    doc.get('LatestPublishedVersionId', '')
                ]
                writer.writerow(row)
        
        self._log_status(f"✅ CSV file created: {output_path}")
        self._log_status(f"✅ Total records exported: {len(content_documents)}")
        return output_path
    
    def _log_status(self, message: str):
        """Log status message"""
        if self.sf_client.status_callback:
            self.sf_client.status_callback(message, verbose=True)
