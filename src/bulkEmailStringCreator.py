import re

def create_bulk_email_string(emails):
    # Extract all potential email addresses using a regular expression
    email_pattern = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+'
    potential_emails = re.findall(email_pattern, emails)

    # Validate and correct the emails
    corrected_email_list = []
    for email in potential_emails:
        email = email.strip().lower()
        
        # Get the part after the @ symbol
        domain_part = email.split('@')[1]
        
        # If the domain part doesn't contain a dot, or ends with specific patterns, correct it
        if '.' not in domain_part:
            email += '.com'
        elif domain_part.endswith('.c'):
            email += 'om'
        elif domain_part.endswith('.co'):
            email += 'm'
        elif domain_part.endswith('.o'):
            email += 'rg'
        elif domain_part.endswith('.or'):
            email += 'g'
        elif domain_part.endswith('.n'):
            email += 'et'
        elif domain_part.endswith('.ne'):
            email += 't'
        elif domain_part.endswith('.g'):
            email += 'ov'
        elif domain_part.endswith('.go'):
            email += 'v'
        elif domain_part.endswith('.e'):
            email += 'du'
        elif domain_part.endswith('.ed'):
            email += 'u'
        elif domain_part.endswith('.i'):
            email += 'nfo'
        elif domain_part.endswith('.in'):
            email += 'fo'
        elif domain_part.endswith('.m'):
            email += 'il'
        elif domain_part.endswith('.mi'):
            email += 'l'
        elif domain_part.endswith('.t'):
            email += 'nt'
        elif domain_part.endswith('.in'):
            email += 't'
            
        # Use a regex pattern to validate the corrected email
        match = re.match(r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$', email)
        if match:
            corrected_email_list.append(email)

    # Add the three extra email addresses
    extra_emails = ['lazaro.gonzalez@conduent.com', 'mary.rodriguez@conduent.com']
    corrected_email_list.extend(extra_emails)
    
    # Remove duplicates
    corrected_email_list = list(set(corrected_email_list))

    bcc_string = '; '.join(corrected_email_list)
    return bcc_string


