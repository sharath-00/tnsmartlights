"""
Schnell Energy Equipments - Universal Smart Lights Dashboard Generator
Compatibility wrapper redirecting to generate_dashboard.py
"""
import sys
from generate_dashboard import run_common_generator

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Schnell Smart Lights Universal Dashboard Generator")
    parser.add_argument("--project", default="all", help="Project key ('all', 'trichy', 'vellore', etc.)")
    parser.add_argument("--send", action="store_true", help="Send email report via SMTP")
    parser.add_argument("--to", default=None, help="Comma-separated recipient email addresses")
    parser.add_argument("--cached", action="store_true", help="Use saved CSV cache instead of live portal scraping")

    args = parser.parse_args()
    run_common_generator(project=args.project, send_email=args.send, custom_recipients=args.to, live=not args.cached)
