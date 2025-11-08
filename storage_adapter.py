"""
Storage adapter for presentation persistence.
Supports multiple backends: S3, Render Disk, and local filesystem.
"""

import os
from typing import Optional, BinaryIO
import tempfile
import shutil


class StorageAdapter:
    """Abstract storage adapter for presentation persistence."""
    
    def __init__(self):
        # Default to 'disk' for Render persistent storage (no external setup needed)
        self.storage_type = os.getenv('STORAGE_TYPE', 'disk').lower()
        self.base_url = os.getenv('BASE_URL', '')
        
        if self.storage_type == 's3':
            self._init_s3()
        elif self.storage_type == 'disk':
            self._init_disk()
        else:
            self._init_local()
    
    def _init_s3(self):
        """Initialize S3 storage."""
        try:
            import boto3
            from botocore.exceptions import ClientError
        except ImportError:
            print("Warning: boto3 not installed. Install with: pip install boto3")
            print("Falling back to local storage")
            self.storage_type = 'local'
            self._init_local()
            return
        
        self.s3_bucket = os.getenv('S3_BUCKET_NAME')
        self.s3_region = os.getenv('S3_REGION', 'us-east-1')
        self.s3_access_key = os.getenv('AWS_ACCESS_KEY_ID')
        self.s3_secret_key = os.getenv('AWS_SECRET_ACCESS_KEY')
        
        if not all([self.s3_bucket, self.s3_access_key, self.s3_secret_key]):
            print("Warning: S3 credentials not fully configured, falling back to local storage")
            self.storage_type = 'local'
            self._init_local()
            return
        
        try:
            self.s3_client = boto3.client(
                's3',
                aws_access_key_id=self.s3_access_key,
                aws_secret_access_key=self.s3_secret_key,
                region_name=self.s3_region
            )
            # Test connection
            self.s3_client.head_bucket(Bucket=self.s3_bucket)
            print(f"S3 storage initialized: bucket={self.s3_bucket}, region={self.s3_region}")
        except Exception as e:
            print(f"Warning: S3 initialization failed: {e}, falling back to local storage")
            self.storage_type = 'local'
            self._init_local()
    
    def _init_disk(self):
        """Initialize Render Disk storage."""
        # Render mounts persistent disks at /mnt/disk by default
        disk_path = os.getenv('DISK_PATH', '/mnt/disk/presentations')
        try:
            os.makedirs(disk_path, exist_ok=True)
            self.disk_path = disk_path
            print(f"Render Disk storage initialized: path={disk_path}")
        except PermissionError:
            # If /mnt/disk doesn't exist (no disk attached), fall back to local
            print("Warning: Render Disk not available, falling back to local storage")
            print("To enable persistent storage, attach a disk in Render dashboard")
            self.storage_type = 'local'
            self._init_local()
    
    def _init_local(self):
        """Initialize local filesystem storage."""
        local_path = os.getenv('PRESENTATIONS_DIR', './presentations')
        os.makedirs(local_path, exist_ok=True)
        self.local_path = local_path
        print(f"Local storage initialized: path={local_path}")
    
    def get_presentation_path(self, filename: str) -> str:
        """Get the full path for a presentation based on storage type."""
        if self.storage_type == 's3':
            return f"s3://{self.s3_bucket}/{filename}"
        elif self.storage_type == 'disk':
            return os.path.join(self.disk_path, filename)
        else:
            return os.path.join(self.local_path, filename)
    
    def download_presentation(self, filename: str, local_path: Optional[str] = None) -> str:
        """
        Download a presentation from storage to local filesystem.
        Returns the local file path.
        """
        if self.storage_type == 's3':
            if local_path is None:
                local_path = os.path.join(tempfile.gettempdir(), filename)
            
            try:
                from botocore.exceptions import ClientError
                self.s3_client.download_file(self.s3_bucket, filename, local_path)
                return local_path
            except ClientError as e:
                if e.response['Error']['Code'] == '404':
                    raise FileNotFoundError(f"Presentation {filename} not found in S3")
                raise
        
        elif self.storage_type == 'disk':
            source_path = os.path.join(self.disk_path, filename)
            if local_path is None:
                local_path = os.path.join(tempfile.gettempdir(), filename)
            
            if not os.path.exists(source_path):
                raise FileNotFoundError(f"Presentation {filename} not found")
            
            shutil.copy2(source_path, local_path)
            return local_path
        
        else:  # local
            file_path = os.path.join(self.local_path, filename)
            if not os.path.exists(file_path):
                raise FileNotFoundError(f"Presentation {filename} not found")
            return file_path
    
    def upload_presentation(self, local_path: str, filename: str) -> str:
        """
        Upload a presentation from local filesystem to storage.
        Returns the storage URL/path.
        """
        if self.storage_type == 's3':
            try:
                from botocore.exceptions import ClientError
                self.s3_client.upload_file(local_path, self.s3_bucket, filename)
                # Return public URL if bucket is public, otherwise return S3 path
                if self.base_url:
                    return f"{self.base_url}/presentations/{filename}"
                else:
                    return f"s3://{self.s3_bucket}/{filename}"
            except ClientError as e:
                raise Exception(f"Failed to upload to S3: {str(e)}")
        
        elif self.storage_type == 'disk':
            dest_path = os.path.join(self.disk_path, filename)
            shutil.copy2(local_path, dest_path)
            if self.base_url:
                return f"{self.base_url}/presentations/{filename}"
            return dest_path
        
        else:  # local
            dest_path = os.path.join(self.local_path, filename)
            shutil.copy2(local_path, dest_path)
            if self.base_url:
                return f"{self.base_url}/presentations/{filename}"
            return dest_path
    
    def presentation_exists(self, filename: str) -> bool:
        """Check if a presentation exists in storage."""
        if self.storage_type == 's3':
            try:
                from botocore.exceptions import ClientError
                self.s3_client.head_object(Bucket=self.s3_bucket, Key=filename)
                return True
            except ClientError as e:
                if e.response['Error']['Code'] == '404':
                    return False
                raise
        
        elif self.storage_type == 'disk':
            return os.path.exists(os.path.join(self.disk_path, filename))
        
        else:  # local
            return os.path.exists(os.path.join(self.local_path, filename))
    
    def delete_presentation(self, filename: str) -> bool:
        """Delete a presentation from storage."""
        if self.storage_type == 's3':
            try:
                from botocore.exceptions import ClientError
                self.s3_client.delete_object(Bucket=self.s3_bucket, Key=filename)
                return True
            except ClientError as e:
                raise Exception(f"Failed to delete from S3: {str(e)}")
        
        elif self.storage_type == 'disk':
            file_path = os.path.join(self.disk_path, filename)
            if os.path.exists(file_path):
                os.remove(file_path)
                return True
            return False
        
        else:  # local
            file_path = os.path.join(self.local_path, filename)
            if os.path.exists(file_path):
                os.remove(file_path)
                return True
            return False
    
    def get_presentation_url(self, filename: str) -> str:
        """Get the public URL for a presentation."""
        if self.storage_type == 's3':
            # Generate presigned URL (valid for 1 hour) or use public URL
            try:
                url = self.s3_client.generate_presigned_url(
                    'get_object',
                    Params={'Bucket': self.s3_bucket, 'Key': filename},
                    ExpiresIn=3600
                )
                return url
            except:
                # Fallback to public URL if bucket is public
                return f"https://{self.s3_bucket}.s3.{self.s3_region}.amazonaws.com/{filename}"
        
        elif self.storage_type == 'disk' or self.storage_type == 'local':
            if self.base_url:
                return f"{self.base_url}/presentations/{filename}"
            return self.get_presentation_path(filename)
        
        return filename


# Global storage adapter instance
_storage_adapter: Optional[StorageAdapter] = None


def get_storage_adapter() -> StorageAdapter:
    """Get or create the global storage adapter instance."""
    global _storage_adapter
    if _storage_adapter is None:
        _storage_adapter = StorageAdapter()
    return _storage_adapter

