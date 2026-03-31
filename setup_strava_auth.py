"""
Strava OAuth Setup Script - COMPLETE
One-time setup to obtain access and refresh tokens
Run this once, then use sync_strava_to_excel.py for regular syncing
"""

import os
import sys
import webbrowser
import urllib3
from dotenv import load_dotenv, set_key

# ============================================================================
# CORPORATE NETWORK SSL BYPASS WORKAROUND
# ============================================================================
# ⚠️ WARNING: This section disables SSL certificate verification!
#
# WHY THIS EXISTS:
# Corporate networks often use "man-in-the-middle" SSL inspection with
# self-signed certificates. Python's requests library correctly rejects these
# as security threats, causing errors like:
#   "certificate verify failed: self-signed certificate in certificate chain"
#
# WHAT THIS DOES:
# 1. Disables urllib3 SSL warnings (cosmetic - suppresses console noise)
# 2. Monkey-patches the requests library to force verify=False on ALL HTTPS
#    requests made by stravalib internally
#
# FOR HOME NETWORKS / PERSONAL COMPUTERS:
# If you're on a home WiFi network WITHOUT corporate SSL inspection, you can
# COMMENT OUT or DELETE lines 13-22 below (everything between the === bars).
# The script will work normally WITH proper SSL certificate verification.
#
# SECURITY NOTE:
# Disabling SSL verification means your connection to Strava API is vulnerable
# to man-in-the-middle attacks. Only use this on TRUSTED corporate networks.
# NEVER use this on public WiFi.
# ============================================================================

# Step 1: Suppress SSL warning messages (cosmetic only)
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

import requests

# Step 2: Globally disable SSL certificate verification for all requests
# This monkey-patches the requests.Session class used internally by stravalib
original_request = requests.Session.request
def patched_request(self, *args, **kwargs):
    kwargs['verify'] = False  # Force SSL verification OFF
    return original_request(self, *args, **kwargs)
requests.Session.request = patched_request

# ============================================================================
# END CORPORATE NETWORK WORKAROUND - Safe to import stravalib now
# ============================================================================

from stravalib.client import Client

def main():
    """Main OAuth setup workflow"""
    
    print("=" * 60)
    print("🚴 STRAVA OAUTH SETUP")
    print("=" * 60)
    print()
    
    # Load environment variables
    print("📂 Loading .env file...")
    env_file = '.env'
    load_dotenv(env_file)
    
    # Get credentials
    client_id = os.getenv('CLIENT_ID')
    client_secret = os.getenv('CLIENT_SECRET')
    
    # Validate credentials exist
    if not client_id or not client_secret:
        print("❌ ERROR: CLIENT_ID or CLIENT_SECRET not found in .env file")
        print()
        print("Please:")
        print("1. Copy .env.template to .env")
        print("2. Fill in your CLIENT_ID and CLIENT_SECRET from Strava API page")
        print("3. Re-run this script")
        sys.exit(1)
    
    print(f"✅ CLIENT_ID found: {client_id}")
    print(f"✅ CLIENT_SECRET found: {'*' * len(client_secret)}")
    print()
    
    # Initialize Strava client
    print("🔧 Initializing Strava client...")
    client = Client()
    print("✅ Strava client ready")
    print()
    
    # Generate authorization URL
    print("=" * 60)
    print("STEP 1: AUTHORIZE APPLICATION")
    print("=" * 60)
    print()
    
    authorize_url = client.authorization_url(
        client_id=client_id,
        redirect_uri='http://localhost',
        scope=['read', 'activity:read_all']
    )
    
    print("📋 Authorization URL generated")
    print()
    print("🌐 Opening your browser to Strava authorization page...")
    print()
    print("What will happen:")
    print("1. Browser opens to Strava")
    print("2. You click 'Authorize' to grant access")
    print("3. Browser redirects to localhost (will show 'can't connect')")
    print("4. Copy the FULL URL from your browser address bar")
    print()
    
    # Open browser
    try:
        webbrowser.open(authorize_url)
        print("✅ Browser opened")
    except Exception as e:
        print(f"⚠️  Could not auto-open browser: {e}")
        print()
        print("Please manually open this URL:")
        print(authorize_url)
    
    print()
    print("-" * 60)
    print("After clicking 'Authorize' on Strava...")
    print("-" * 60)
    print()
    
    # Get authorization code from user
    print("📋 The redirect URL will look like:")
    print("http://localhost/?state=&code=LONG_CODE_HERE&scope=read,activity:read_all")
    print()
    
    redirect_url = input("Paste the FULL redirect URL here and press Enter: ").strip()
    
    if not redirect_url:
        print("❌ ERROR: No URL provided")
        sys.exit(1)
    
    # Extract authorization code
    print()
    print("🔍 Extracting authorization code...")
    
    try:
        # Parse code from URL
        if 'code=' not in redirect_url:
            print("❌ ERROR: No authorization code found in URL")
            print("Make sure you copied the FULL URL after clicking Authorize")
            sys.exit(1)
        
        code = redirect_url.split('code=')[1].split('&')[0]
        print(f"✅ Authorization code extracted: {code[:20]}...")
        
    except Exception as e:
        print(f"❌ ERROR parsing URL: {e}")
        print("Make sure you copied the complete redirect URL")
        sys.exit(1)
    
    # Exchange code for tokens
    print()
    print("=" * 60)
    print("STEP 2: EXCHANGE CODE FOR TOKENS")
    print("=" * 60)
    print()
    print("🔄 Exchanging authorization code for access token...")
    
    try:
        token_response = client.exchange_code_for_token(
            client_id=client_id,
            client_secret=client_secret,
            code=code
        )
        
        access_token = token_response['access_token']
        refresh_token = token_response['refresh_token']
        expires_at = token_response['expires_at']
        
        print("✅ Tokens received!")
        print(f"   Access Token: {access_token[:20]}...")
        print(f"   Refresh Token: {refresh_token[:20]}...")
        print(f"   Expires at: {expires_at}")
        
    except Exception as e:
        print(f"❌ ERROR exchanging code for token: {e}")
        print()
        print("Common causes:")
        print("- Authorization code already used (run script again)")
        print("- Code expired (codes expire after a few minutes)")
        print("- CLIENT_ID or CLIENT_SECRET incorrect")
        sys.exit(1)
    
    # Save tokens to .env file
    print()
    print("=" * 60)
    print("STEP 3: SAVE TOKENS TO .ENV FILE")
    print("=" * 60)
    print()
    print("💾 Updating .env file with tokens...")
    
    try:
        # Update .env file with tokens
        set_key(env_file, 'ACCESS_TOKEN', access_token)
        set_key(env_file, 'REFRESH_TOKEN', refresh_token)
        
        print("✅ .env file updated successfully!")
        print()
        print("Your .env file now contains:")
        print("  ✓ CLIENT_ID")
        print("  ✓ CLIENT_SECRET")
        print("  ✓ ACCESS_TOKEN")
        print("  ✓ REFRESH_TOKEN")
        
    except Exception as e:
        print(f"❌ ERROR updating .env file: {e}")
        print()
        print("Manual fallback - add these to your .env file:")
        print(f"ACCESS_TOKEN={access_token}")
        print(f"REFRESH_TOKEN={refresh_token}")
        sys.exit(1)
    
    # Verify token works
    print()
    print("=" * 60)
    print("STEP 4: VERIFY TOKEN")
    print("=" * 60)
    print()
    print("🧪 Testing access token by fetching athlete info...")
    
    try:
        client.access_token = access_token
        athlete = client.get_athlete()
        
        print("✅ Token verified! Successfully connected to your Strava account:")
        print()
        print(f"   👤 Name: {athlete.firstname} {athlete.lastname}")
        print(f"   📍 Location: {athlete.city}, {athlete.state}, {athlete.country}")
        print(f"   🎽 Premium: {'Yes' if athlete.premium else 'No'}")
        
    except Exception as e:
        print(f"⚠️  Token saved but verification failed: {e}")
        print("This might be okay - try running sync_strava_to_excel.py")
    
    # Success summary
    print()
    print("=" * 60)
    print("🎉 SETUP COMPLETE!")
    print("=" * 60)
    print()
    print("✅ OAuth setup successful")
    print("✅ Tokens saved to .env file")
    print("✅ Ready to sync Strava activities")
    print()
    print("NEXT STEPS:")
    print("1. Run: python sync_strava_to_excel.py")
    print("2. Your Strava activities will be imported to the workbook")
    print("3. Use the sync script every Monday for weekly updates")
    print()
    print("💡 TIP: Tokens expire after 6 hours, but the script will")
    print("   auto-refresh them using your refresh token")
    print()

if __name__ == '__main__':
    main()
