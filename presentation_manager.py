"""
Presentation manager that handles storage operations transparently.
Downloads presentations before editing, uploads after saving.
"""

import os
import tempfile
from typing import Optional
from storage_adapter import get_storage_adapter


class PresentationManager:
    """Manages presentation lifecycle with automatic storage sync."""
    
    def __init__(self):
        self.storage = get_storage_adapter()
        self.temp_dir = tempfile.mkdtemp(prefix='ppt_edit_')
    
    def get_local_path(self, filename: str, create_if_missing: bool = False) -> str:
        """
        Get local file path for editing.
        Downloads from storage if needed.
        
        Args:
            filename: Presentation filename
            create_if_missing: If True, create empty file if it doesn't exist
        
        Returns:
            Local file path for editing
        """
        # Check if presentation exists in storage
        if self.storage.presentation_exists(filename):
            # Download to temp location for editing
            local_path = os.path.join(self.temp_dir, filename)
            self.storage.download_presentation(filename, local_path)
            return local_path
        elif create_if_missing:
            # Create new presentation in temp location
            local_path = os.path.join(self.temp_dir, filename)
            # Ensure directory exists
            os.makedirs(os.path.dirname(local_path), exist_ok=True)
            return local_path
        else:
            raise FileNotFoundError(f"Presentation {filename} not found")
    
    def save_presentation(self, local_path: str, filename: str) -> str:
        """
        Save presentation to storage and return URL.
        
        Args:
            local_path: Local file path
            filename: Target filename in storage
        
        Returns:
            Presentation URL
        """
        # Upload to storage
        url = self.storage.upload_presentation(local_path, filename)
        return url
    
    def get_presentation_url(self, filename: str) -> str:
        """Get the public URL for a presentation."""
        return self.storage.get_presentation_url(filename)
    
    def cleanup_temp(self, filename: Optional[str] = None):
        """Clean up temporary files."""
        if filename:
            temp_path = os.path.join(self.temp_dir, filename)
            if os.path.exists(temp_path):
                os.remove(temp_path)
        else:
            # Clean up entire temp directory
            import shutil
            if os.path.exists(self.temp_dir):
                shutil.rmtree(self.temp_dir)


# Global presentation manager instance
_presentation_manager: Optional[PresentationManager] = None


def get_presentation_manager() -> PresentationManager:
    """Get or create the global presentation manager instance."""
    global _presentation_manager
    if _presentation_manager is None:
        _presentation_manager = PresentationManager()
    return _presentation_manager

