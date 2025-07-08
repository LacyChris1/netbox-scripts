"""
NetBox Contact to VCF Export Script
==================================

This script exports contacts from a NetBox ContactGroup to a VCF (vCard) file.
It handles the complete workflow: data extraction, validation, formatting, and file generation.

Requirements:
- NetBox 4.3+
- Place this file in /opt/netbox/netbox/scripts/ (or Docker equivalent)
- Ensure the script has proper permissions

Usage:
- Run through NetBox UI: System > Scripts
- Select contact group and run script
- Download generated VCF file
"""

from extras.scripts import Script, ObjectVar, ChoiceVar, BooleanVar, StringVar
from tenancy.models import Contact, ContactGroup
from django.http import HttpResponse
from django.utils.text import slugify
from django.core.exceptions import ValidationError
import re
import uuid
from datetime import datetime
import logging

# Set up logging for debugging
logger = logging.getLogger(__name__)

class ContactToVCFExport(Script):
    """
    Export contacts from a NetBox ContactGroup to VCF format.
    
    This script demonstrates several key concepts:
    1. Data extraction from NetBox models
    2. Data validation and sanitization
    3. Format transformation (NetBox → VCF)
    4. File generation and download
    """
    
    class Meta:
        name = "Export Contacts to VCF"
        description = "Export all contacts from a selected group to a VCF (vCard) file"
        commit_default = False  # This is a read-only operation
        
    # Script parameters - these create the user interface
    contact_group = ObjectVar(
        model=ContactGroup,
        description="Select the contact group to export",
        required=True
    )
    
    include_subgroups = BooleanVar(
        description="Include contacts from subgroups (recursive)",
        default=True,
        required=False
    )
    
    vcf_version = ChoiceVar(
        choices=[
            ('3.0', 'vCard 3.0 (Most Compatible)'),
            ('4.0', 'vCard 4.0 (Latest Standard)')
        ],
        default='3.0',
        description="VCF format version",
        required=False
    )
    
    filename_prefix = StringVar(
        description="Filename prefix (optional)",
        default="netbox_contacts",
        required=False
    )

    def run(self, data, commit):
        """
        Main execution method - this is where the magic happens!
        
        The 'data' parameter contains the user's form input.
        The 'commit' parameter indicates if changes should be saved (not relevant here).
        """
        
        # Step 1: Extract and validate input parameters
        contact_group = data['contact_group']
        include_subgroups = data.get('include_subgroups', True)
        vcf_version = data.get('vcf_version', '3.0')
        filename_prefix = data.get('filename_prefix', 'netbox_contacts')
        
        self.log_info(f"Starting VCF export for group: {contact_group.name}")
        self.log_info(f"Include subgroups: {include_subgroups}")
        self.log_info(f"VCF Version: {vcf_version}")
        
        # Step 2: Gather all contacts from the selected group
        try:
            contacts = self._gather_contacts(contact_group, include_subgroups)
            self.log_info(f"Found {len(contacts)} contacts to export")
            
            if not contacts:
                self.log_warning("No contacts found in the selected group")
                return "No contacts found to export"
                
        except Exception as e:
            self.log_failure(f"Error gathering contacts: {str(e)}")
            return f"Error gathering contacts: {str(e)}"
        
        # Step 3: Validate and clean contact data
        try:
            validated_contacts = self._validate_contacts(contacts)
            self.log_info(f"Validated {len(validated_contacts)} contacts")
            
        except Exception as e:
            self.log_failure(f"Error validating contacts: {str(e)}")
            return f"Error validating contacts: {str(e)}"
        
        # Step 4: Generate VCF content
        try:
            vcf_content = self._generate_vcf_content(validated_contacts, vcf_version)
            self.log_success("VCF content generated successfully")
            
        except Exception as e:
            self.log_failure(f"Error generating VCF content: {str(e)}")
            return f"Error generating VCF content: {str(e)}"
        
        # Step 5: Create and save the file
        try:
            filename = self._generate_filename(filename_prefix, contact_group.name)
            self._save_vcf_file(vcf_content, filename)
            
            self.log_success(f"VCF file '{filename}' created successfully")
            return f"Successfully exported {len(validated_contacts)} contacts to {filename}"
            
        except Exception as e:
            self.log_failure(f"Error saving VCF file: {str(e)}")
            return f"Error saving VCF file: {str(e)}"

    def _gather_contacts(self, contact_group, include_subgroups):
        """
        Gather all contacts from the specified group.
        
        This method handles the recursive nature of contact groups and the many-to-many
        relationship between contacts and groups. A contact can belong to multiple groups,
        so we use the 'groups' field (plural) to find contacts.
        """
        contacts = []
        
        # Get direct contacts from this group using the many-to-many relationship
        # We use 'groups' (plural) because contacts can belong to multiple groups
        direct_contacts = Contact.objects.filter(groups=contact_group)
        contacts.extend(direct_contacts)
        
        self.log_info(f"Found {len(direct_contacts)} direct contacts in '{contact_group.name}'")
        
        # If requested, recursively get contacts from subgroups
        if include_subgroups:
            subgroups = ContactGroup.objects.filter(parent=contact_group)
            
            for subgroup in subgroups:
                self.log_info(f"Processing subgroup: {subgroup.name}")
                subgroup_contacts = self._gather_contacts(subgroup, include_subgroups)
                contacts.extend(subgroup_contacts)
        
        # Remove duplicates (in case a contact appears in multiple groups)
        # This is especially important with many-to-many relationships
        unique_contacts = list(set(contacts))
        
        if len(contacts) != len(unique_contacts):
            self.log_info(f"Removed {len(contacts) - len(unique_contacts)} duplicate contacts")
        
        return unique_contacts

    def _validate_contacts(self, contacts):
        """
        Validate and clean contact data before VCF generation.
        
        This step is crucial because NetBox allows flexible data entry,
        but VCF format has specific requirements.
        """
        validated_contacts = []
        
        for contact in contacts:
            try:
                # Create a cleaned contact dictionary
                cleaned_contact = {
                    'id': contact.id,
                    'name': self._clean_name(contact.name),
                    'email': self._clean_email(contact.email),
                    'phone': self._clean_phone(contact.phone),
                    'title': getattr(contact, 'title', ''),
                    'address': getattr(contact, 'address', ''),
                    'comments': getattr(contact, 'comments', ''),
                    'groups': [group.name for group in contact.groups.all()],  # Get all group names
                    'original_contact': contact  # Keep reference for debugging
                }
                
                # Validate that we have at least a name
                if not cleaned_contact['name']:
                    self.log_warning(f"Skipping contact with ID {contact.id}: No name provided")
                    continue
                
                # Validate that we have at least one contact method
                if not cleaned_contact['email'] and not cleaned_contact['phone']:
                    self.log_warning(f"Skipping contact '{cleaned_contact['name']}': No email or phone provided")
                    continue
                
                validated_contacts.append(cleaned_contact)
                
            except Exception as e:
                self.log_warning(f"Error processing contact ID {contact.id}: {str(e)}")
                continue
        
        return validated_contacts

    def _clean_name(self, name):
        """Clean and validate contact name."""
        if not name:
            return ""
        
        # Remove any characters that might cause issues in VCF
        cleaned = re.sub(r'[^\w\s\-\.]', '', str(name).strip())
        return cleaned[:100]  # Limit length to prevent issues

    def _clean_email(self, email):
        """Clean and validate email address."""
        if not email:
            return ""
        
        # Basic email validation regex
        email_pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        email = str(email).strip().lower()
        
        if re.match(email_pattern, email):
            return email
        else:
            return ""  # Invalid email, return empty string

    def _clean_phone(self, phone):
        """Clean and validate phone number."""
        if not phone:
            return ""
        
        # Remove all non-numeric characters except +
        cleaned = re.sub(r'[^\d+]', '', str(phone))
        
        # Basic phone number validation (at least 7 digits)
        if len(re.sub(r'[^\d]', '', cleaned)) >= 7:
            return cleaned
        else:
            return ""  # Invalid phone, return empty string

    def _generate_vcf_content(self, contacts, vcf_version):
        """
        Generate VCF content from validated contacts.
        
        This method creates proper vCard format based on the selected version.
        VCF format is quite specific about line endings and structure.
        """
        vcf_lines = []
        
        for contact in contacts:
            # Start vCard
            vcf_lines.append("BEGIN:VCARD")
            vcf_lines.append(f"VERSION:{vcf_version}")
            
            # Add name (required field)
            # FN = Formatted Name, N = Structured Name
            vcf_lines.append(f"FN:{contact['name']}")
            
            # For structured name, we'll try to split first/last name
            name_parts = contact['name'].split()
            if len(name_parts) >= 2:
                last_name = name_parts[-1]
                first_name = " ".join(name_parts[:-1])
                vcf_lines.append(f"N:{last_name};{first_name};;;")
            else:
                vcf_lines.append(f"N:{contact['name']};;;;")
            
            # Add email if available
            if contact['email']:
                if vcf_version == '4.0':
                    vcf_lines.append(f"EMAIL:{contact['email']}")
                else:
                    vcf_lines.append(f"EMAIL;TYPE=INTERNET:{contact['email']}")
            
            # Add phone if available
            if contact['phone']:
                if vcf_version == '4.0':
                    vcf_lines.append(f"TEL:{contact['phone']}")
                else:
                    vcf_lines.append(f"TEL;TYPE=VOICE:{contact['phone']}")
            
            # Add title if available
            if contact['title']:
                vcf_lines.append(f"TITLE:{contact['title']}")
            
            # Add address if available
            if contact['address']:
                vcf_lines.append(f"ADR:;;{contact['address']};;;;")
            
            # Add notes/comments if available
            notes = []
            if contact['comments']:
                notes.append(f"Comments: {contact['comments']}")
            
            # Add group membership information
            if contact['groups']:
                group_list = ", ".join(contact['groups'])
                notes.append(f"Groups: {group_list}")
            
            if notes:
                note_text = " | ".join(notes)
                vcf_lines.append(f"NOTE:{note_text}")
            
            # Add unique ID
            vcf_lines.append(f"UID:{uuid.uuid4()}")
            
            # Add revision timestamp
            timestamp = datetime.now().strftime("%Y%m%dT%H%M%SZ")
            vcf_lines.append(f"REV:{timestamp}")
            
            # End vCard
            vcf_lines.append("END:VCARD")
            vcf_lines.append("")  # Empty line between contacts
        
        # Join all lines with proper line endings
        return "\r\n".join(vcf_lines)

    def _generate_filename(self, prefix, group_name):
        """Generate a safe filename for the VCF file."""
        # Clean the group name for use in filename
        safe_group_name = slugify(group_name)
        
        # Create timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # Combine into filename
        filename = f"{prefix}_{safe_group_name}_{timestamp}.vcf"
        
        return filename

    def _save_vcf_file(self, content, filename):
        """
        Save VCF content to a file.
        
        In NetBox, custom scripts can write to the media directory,
        which is typically accessible via the web interface.
        """
        import os
        from django.conf import settings
        
        # Define the output directory
        output_dir = os.path.join(settings.MEDIA_ROOT, 'vcf_exports')
        
        # Create directory if it doesn't exist
        os.makedirs(output_dir, exist_ok=True)
        
        # Full file path
        file_path = os.path.join(output_dir, filename)
        
        # Write the file
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(content)
        
        # Log the file location
        self.log_info(f"VCF file saved to: {file_path}")
        
        # Make the file accessible via web interface
        # The media URL will be something like: /media/vcf_exports/filename.vcf
        media_url = f"{settings.MEDIA_URL}vcf_exports/{filename}"
        self.log_info(f"File accessible at: {media_url}")


# Additional utility script for direct API usage
class ContactVCFExportAPI:
    """
    Alternative approach using NetBox's REST API directly.
    
    This class can be used if you prefer to interact with NetBox
    via its REST API rather than using custom scripts.
    """
    
    def __init__(self, netbox_url, api_token):
        """
        Initialize the API client.
        
        Args:
            netbox_url (str): NetBox instance URL (e.g., 'https://netbox.example.com')
            api_token (str): NetBox API token
        """
        self.netbox_url = netbox_url.rstrip('/')
        self.api_token = api_token
        self.headers = {
            'Authorization': f'Token {api_token}',
            'Content-Type': 'application/json'
        }
    
    def get_contact_groups(self):
        """Retrieve all contact groups."""
        import requests
        
        url = f"{self.netbox_url}/api/tenancy/contact-groups/"
        response = requests.get(url, headers=self.headers)
        response.raise_for_status()
        
        return response.json()['results']
    
    def get_contacts_by_group(self, group_id):
        """
        Retrieve all contacts for a specific group.
        
        Note: This uses the 'groups' field (plural) because contacts can belong
        to multiple groups simultaneously in NetBox's many-to-many relationship model.
        """
        import requests
        
        url = f"{self.netbox_url}/api/tenancy/contacts/"
        params = {'groups': group_id}  # Changed from 'group_id' to 'groups'
        
        response = requests.get(url, headers=self.headers, params=params)
        response.raise_for_status()
        
        return response.json()['results']
    
    def export_group_to_vcf(self, group_id, filename=None):
        """Export a contact group to VCF format."""
        contacts = self.get_contacts_by_group(group_id)
        
        if not contacts:
            raise ValueError("No contacts found for the specified group")
        
        # Generate VCF content (simplified version)
        vcf_content = self._generate_simple_vcf(contacts)
        
        if filename:
            with open(filename, 'w', encoding='utf-8') as f:
                f.write(vcf_content)
        
        return vcf_content
    
    def _generate_simple_vcf(self, contacts):
        """Generate simple VCF content from API contact data."""
        vcf_lines = []
        
        for contact in contacts:
            vcf_lines.append("BEGIN:VCARD")
            vcf_lines.append("VERSION:3.0")
            vcf_lines.append(f"FN:{contact['name']}")
            vcf_lines.append(f"N:{contact['name']};;;;")
            
            if contact.get('email'):
                vcf_lines.append(f"EMAIL;TYPE=INTERNET:{contact['email']}")
            
            if contact.get('phone'):
                vcf_lines.append(f"TEL;TYPE=VOICE:{contact['phone']}")
            
            vcf_lines.append(f"UID:{uuid.uuid4()}")
            vcf_lines.append("END:VCARD")
            vcf_lines.append("")
        
        return "\r\n".join(vcf_lines)


# Example usage and testing
if __name__ == "__main__":
    """
    This section would only run if the script is executed directly,
    not when imported as a NetBox custom script.
    
    Use this for testing your logic outside of NetBox.
    """
    print("This script is designed to run as a NetBox custom script.")
    print("Place it in /opt/netbox/netbox/scripts/ and run it through the NetBox UI.")
    
    # Example of how to use the API class:
    # api = ContactVCFExportAPI('https://your-netbox.com', 'your-api-token')
    # groups = api.get_contact_groups()
    # print(f"Found {len(groups)} contact groups")
