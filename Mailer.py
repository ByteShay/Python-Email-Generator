import pandas as pd
from pathlib import Path

def generate_email(business, email_address, has_website):
    business_name = business['name']

    # Updated signature — simpler and less promotional
    signature = """

Shay M.  
Lead Developer  
Shay's Web Services  
"""

    if has_website:
        email_content = f"""To: {email_address}
Subject: Suggestions to Improve Your Website

Hi {business_name} Team,

I recently reviewed your website and noticed a few areas where small updates could make a strong impact — particularly with mobile performance and visibility on search engines.

My team supports small businesses by helping with:

- Clean, modern design updates  
- Faster load speeds and responsive layouts  
- Search engine improvements  

I’d be happy to share a few suggestions tailored to your site — no commitment necessary.

Best regards,"""
    else:
        email_content = f"""To: {email_address}
Subject: Helping You Get Online with a Simple Website

Hi {business_name} Team,

I wasn’t able to find a website for your business, so I thought I’d reach out. A basic website can help customers find you and learn more about your services.

We work with small businesses to create:

- Simple, mobile-friendly websites  
- Pages that are easy to update  
- Professional layouts that match your brand  

If you'd like to explore this, I’d be happy to chat.

Best regards,"""

    return email_content + signature

def main():
    print("Starting email generator...")
    print("This will create professional emails for all businesses in your list.")

    email_folder = Path('generated_emails')
    email_folder.mkdir(exist_ok=True)

    businesses_processed = 0
    emails_created = 0

    try:
        business_list = pd.read_excel('businesses.xlsx')
        print(f"\nFound {len(business_list)} businesses to process...")

        for _, business in business_list.iterrows():
            businesses_processed += 1
            business_name = business['name']

            for email_field in ['email_1', 'email_2', 'email_3']:
                business_email = business[email_field]

                if pd.notna(business_email) and '@' in str(business_email):
                    has_website = pd.notna(business['site']) and '.' in str(business['site'])

                    email_text = generate_email(business, business_email, has_website)

                    safe_name = business_name.replace(' ', '_')[:30]
                    filename = email_folder / f"Email_to_{safe_name}.txt"

                    with open(filename, 'w', encoding='utf-8') as f:
                        f.write(email_text)

                    emails_created += 1
                    print(f"Created email for {business_name}")
                    break

        print("\nFINISHED!")
        print(f"Processed {businesses_processed} businesses")
        print(f"Created {emails_created} professional emails")
        print(f"Your emails are ready in the '{email_folder}' folder")
        print("You can now open these files, review them, and send to clients!")

    except FileNotFoundError:
        print("\nERROR: Couldn't find 'businesses.xlsx'")
        print("Please make sure your business list has this exact name")
    except Exception as e:
        print(f"\nERROR: Something went wrong - {str(e)}")
        print("Please check your Excel file and try again")

if __name__ == "__main__":
    main()
