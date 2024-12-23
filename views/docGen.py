import streamlit as st
from docx import Document as DocxDocument  # For Word document handling
import tempfile
import re  # For regex operations

# Document templates
templates = {
    "bail_application": """\
[COURT NAME]
[COURT ADDRESS]

[Bail Application No. XYZ]
Date: [DATE]

To,
The Hon'ble Judge,
[COURT NAME]

Re: Bail Application for [APPLICANT NAME] (Case No. [CASE NUMBER])

Dear Sir/Madam,

I, [YOUR NAME], am writing to you in my capacity as the [relation to applicant, e.g., family member, friend, legal counsel] of [APPLICANT NAME], who is currently charged with [OFFENSE] under [SECTION] of the Indian Penal Code (IPC). The charges stem from an incident that occurred on [DATE OF INCIDENT], and I wish to present this application for bail on the following grounds:

1. **Nature of the Offense**: [Briefly describe the nature of the offense and any mitigating circumstances that may apply. If the offense is bailable under IPC, specify it.]

2. **Non-flight Risk**: [Explain why the applicant is not a flight risk. Include details such as employment status, family ties, and community ties.]

3. **Cooperation with Investigation**: [Highlight any cooperation the applicant has shown during the investigation, such as attending police summons or providing statements.]

4. **Health and Well-being**: [Mention any health issues the applicant may have that necessitate their release on bail, if applicable.]

5. **Grounds for Bail**: [List any other pertinent factors that support the request for bail, such as lack of prior convictions, good character references, etc.]

In light of these considerations, I respectfully urge your Honour to grant bail to [APPLICANT NAME]. I assure you that the applicant will adhere to all conditions set forth by the court.

Thank you for your kind consideration.

Sincerely,
[YOUR NAME]
[YOUR ADDRESS]
[YOUR CONTACT]


""",

    "lease_agreement": """\
LEASE AGREEMENT

This Lease Agreement ("Agreement") is made and entered into on this [DATE] by and between [LANDLORD NAME], residing at [LANDLORD ADDRESS] (hereinafter referred to as the "Landlord"), and [TENANT NAME], residing at [TENANT ADDRESS] (hereinafter referred to as the "Tenant").

WHEREAS, the Landlord is the lawful owner of the residential property located at [PROPERTY ADDRESS] (hereinafter referred to as the "Property"), and

WHEREAS, the Tenant wishes to lease the Property from the Landlord for residential purposes,

NOW, THEREFORE, in consideration of the mutual promises contained herein, the parties agree as follows:

1. **Property Address**: The Landlord hereby leases to the Tenant the Property situated at [PROPERTY ADDRESS].

2. **Lease Term**: This lease shall commence on [START DATE] and shall continue until [END DATE] unless terminated earlier in accordance with this Agreement.

3. **Monthly Rent**: The Tenant agrees to pay a monthly rent of ₹[MONTHLY RENT], payable in advance on or before the [DUE DATE] of each month.

4. **Security Deposit**: The Tenant shall pay a security deposit of ₹[SECURITY DEPOSIT] prior to taking possession of the Property. This deposit shall be refunded upon termination of this lease, subject to any deductions for damages or unpaid dues.

5. **Maintenance Responsibilities**: 
   - The Landlord shall be responsible for the maintenance of the Property, including repairs to plumbing, electrical, and structural issues.
   - The Tenant shall maintain the Property in a clean and sanitary condition and shall be responsible for minor repairs.

6. **Late Payment Penalty**: Any rent not received within [NUMBER OF DAYS] days of the due date shall incur a late fee of ₹[LATE PAYMENT PENALTY].

7. **Subletting Clause**: The Tenant shall not sublet the Property or assign this Agreement without the prior written consent of the Landlord.

8. **Termination**: Either party may terminate this Agreement by providing [NUMBER OF DAYS] days' written notice to the other party.

9. **Governing Law**: This Agreement shall be governed by and construed in accordance with the laws of India.

IN WITNESS WHEREOF, the parties have executed this Lease Agreement as of the date first above written.

**Signatures**:
Landlord: ________________  Date: __________
Tenant: ________________    Date: __________
Witness 1: ________________  Date: __________
Witness 2: ________________  Date: __________


""",

    "cease_and_desist": """\
[YOUR LAW FIRM NAME]
[YOUR LAW FIRM ADDRESS]
[PHONE NUMBER]
[EMAIL]

Date: [DATE]

To,
[RECIPIENT NAME]
[RECIPIENT ADDRESS]

Subject: Cease and Desist for Copyright Infringement

Dear [RECIPIENT NAME],

We represent [YOUR CLIENT NAME], the owner of certain copyrighted material, including [DESCRIPTION OF COPYRIGHTED MATERIAL]. It has come to our attention that you have been engaging in activities that infringe upon our client's copyright, specifically [DESCRIPTION OF INFRINGEMENT].

As you are likely aware, such unauthorized use constitutes a violation of the Copyright Act, 1957, and may expose you to legal liability.

We hereby demand that you:

1. Immediately cease and desist all infringing activities, including but not limited to [SPECIFIC ACTIVITIES].
2. Provide us with written confirmation of your compliance by [DEADLINE].

Failure to comply with this demand may result in legal action against you, including seeking damages and injunctive relief.

We hope to resolve this matter amicably. Please contact us at your earliest convenience to discuss this matter further.

Sincerely,
[YOUR NAME]
[YOUR TITLE]
[YOUR LAW FIRM NAME]
[YOUR CONTACT]


""",

    "power_of_attorney": """\
POWER OF ATTORNEY

This Power of Attorney ("POA") is executed on this [DATE] by [PRINCIPAL NAME], residing at [PRINCIPAL ADDRESS] (hereinafter referred to as the "Principal"), appointing [AGENT NAME], residing at [AGENT ADDRESS], as my attorney-in-fact (hereinafter referred to as the "Agent").

WHEREAS, the Principal desires to appoint the Agent to act on their behalf in legal and financial matters, and

WHEREAS, the Agent agrees to accept such appointment under the terms and conditions set forth in this document,

NOW, THEREFORE, the Principal hereby grants the Agent the following powers:

1. **Authority**: The Agent shall have full power and authority to act on behalf of the Principal in the following matters:
   - Managing bank accounts, including deposits and withdrawals.
   - Buying, selling, and managing real estate.
   - Handling legal matters, including but not limited to initiating or defending lawsuits.
   - Making healthcare decisions in accordance with the Principal's wishes.

2. **Effective Date**: This Power of Attorney shall be effective immediately upon execution and shall continue until revoked by the Principal.

3. **Revocation**: The Principal may revoke this Power of Attorney at any time by providing written notice to the Agent. Such revocation shall not affect any actions taken by the Agent prior to receipt of the revocation notice.

4. **Indemnification**: The Principal agrees to indemnify and hold harmless the Agent from any liability arising from actions taken in good faith under this Power of Attorney.

5. **Governing Law**: This POA shall be governed by and construed in accordance with the laws of India.

IN WITNESS WHEREOF, the Principal has executed this Power of Attorney as of the date first above written.

**Signatures**:
Principal: ________________  Date: __________
Agent: ________________      Date: __________
Witness 1: ________________  Date: __________
Witness 2: ________________  Date: __________

"""
}

# Legal Document Generator
def generate_legal_document(prompt):
    # Determine the document type from the prompt
    if "bail application" in prompt.lower():
        template = templates["bail_application"]
        file_name = "bail application"

        # Extract details from prompt
        applicant_name = re.search(r"(?<=for )[\w\s]+", prompt)
        applicant_name = applicant_name.group(0) if applicant_name else "Default Applicant"

        data = {
            "court_name": "District Court of New Delhi",
            "court_address": "123 Court St, New Delhi",
            "applicant_name": applicant_name,
            "parent_name": "Raj Singh",
            "age": "30",
            "address": "45 Elm St, New Delhi",
            "case_number": "A1234",
            "offense": "Theft under Section 379 of IPC",
            "offense_date": "October 1, 2023",
            "grounds_for_bail": "The applicant is a first-time offender with a permanent address and no flight risk.",
            "date": "November 2, 2023"
        }

    elif "lease agreement" in prompt.lower():
        template = templates["lease_agreement"]
        file_name = "lease agreement"

        property_address = re.search(r"(?<=address: )[\w\s,]+", prompt)
        property_address = property_address.group(0) if property_address else "123 Default St, City"

        monthly_rent = re.search(r"(?<=rent: )\d+", prompt)
        monthly_rent = monthly_rent.group(0) if monthly_rent else "1000"

        data = {
            "property_address": property_address,
            "lease_term": "12 months",
            "monthly_rent": f"${monthly_rent}",
            "security_deposit": "$3000",
            "landlord_name": "John Doe",
            "landlord_address": "678 Maple Ave, Mumbai",
            "tenant_name": "Jane Smith",
            "tenant_address": "45 Oak St, Mumbai",
            "maintenance_responsibilities": "Tenant is responsible for minor repairs; landlord for major repairs.",
            "late_payment_penalty": "$50 after 5 days of delay",
            "subletting_clause": "No subletting without landlord's consent",
            "landlord_signature_date": "November 2, 2023",
            "tenant_signature_date": "November 2, 2023"
        }

    elif "cease and desist" in prompt.lower():
        template = templates["cease_and_desist"]
        file_name = "cease and desist"

        infringement_description = re.search(r"(?<=for )[\w\s]+", prompt)
        infringement_description = infringement_description.group(0) if infringement_description else "Unauthorized use of materials."

        data = {
            "sender_name": "ABC Law Firm",
            "sender_address": "101 Legal Lane, Delhi",
            "recipient_name": "XYZ Corporation",
            "recipient_address": "456 Business Rd, Mumbai",
            "infringement_description": infringement_description,
            "copyright_number": "TX0001234567",
            "notice_date": "November 2, 2023",
            "removal_period": "7 days",
            "sender_signature_date": "November 2, 2023"
        }

    elif "power of attorney" in prompt.lower():
        template = templates["power_of_attorney"]
        file_name = "power of attorney"

        principal_name = re.search(r"(?<=for )[\w\s]+", prompt)
        principal_name = principal_name.group(0) if principal_name else "Default Principal"

        data = {
            "principal_name": principal_name,
            "principal_address": "789 Willow Ave, Pune",
            "agent_name": "Michael Blue",
            "agent_address": "321 Cedar St, Pune",
            "property_address": "456 Maple St, Pune",
            "authority_description": "Manage and oversee rental operations for the property.",
            "start_date": "November 2, 2023",
            "end_date": "November 2, 2024",
            "revocation_clause": "This power of attorney may be revoked at any time by the principal.",
            "principal_signature_date": "November 2, 2023",
            "witness_signature_date": "November 2, 2023"
        }
    
    else:
        st.error("Document type not recognized. Please specify a valid document type in your prompt.")
        return

    # Fill in the template with data
    document_content = template.format(**data)

    # Create a Word document with generated content
    doc = DocxDocument()
    for line in document_content.split("\n"):
        if line.strip():  # Avoid adding empty lines
            doc.add_paragraph(line)

    # Save the document to a temporary file and provide it for download in Streamlit
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        doc.save(tmp.name)

        # Read the file to ensure compatibility with Streamlit download
        with open(tmp.name, "rb") as file:
            st.download_button(
                label="Download Generated Document",
                data=file,
                file_name=file_name+".docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

# Legal Document Generator Section
st.markdown("<h1 style='text-align: center;'>⚖️ AI Legal Assistant ⚖️</h1>", unsafe_allow_html=True)
st.write(" ")
st.subheader("📃 Legal Document Generator", divider="grey")
st.write(" ")
doc_prompt = st.text_input("Enter document request (e.g., 'Bail application for Kumar')")
if doc_prompt:
    generate_legal_document(doc_prompt)
