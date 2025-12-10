"""
Example usage of the Slide Generation API

This script demonstrates how to use the new slide generation endpoints
to create different types of slides programmatically.
"""

import requests
import json

# Base URL - adjust if running on different host/port
BASE_URL = "http://localhost:8000/api/slides"


def example_1_points_slide():
    """Example: Generate a points slide"""
    print("\n" + "="*60)
    print("EXAMPLE 1: Generate Points Slide")
    print("="*60)
    
    url = f"{BASE_URL}/generate-points-slide"
    
    payload = {
        "template_s3_url": "presentations/template.pptx",
        "slide_data": {
            "slide_number": 2,
            "header": "API Integration Comparison",
            "description": "This document compares the use of Google's API with third-party alternatives for Veo 3 integration, highlighting key differences and considerations.",
            "image_url": "https://via.placeholder.com/600x400/3667B2/FFFFFF?text=Overview+Image",
            "points": [
                {
                    "text": "Integration Approach: Direct API vs Third-Party SDK",
                    "color": "#000000",
                    "font_size": 14
                },
                {
                    "text": "Cost Analysis: Pricing models and scalability",
                    "color": "#3667B2",
                    "font_size": 14
                },
                {
                    "text": "Performance Benchmarks: Speed and reliability metrics",
                    "color": "#000000",
                    "font_size": 14
                }
            ],
            "header_color": "#3667B2",
            "description_color": "#000000"
        },
        "upload_to_s3": True,
        "output_filename": "points_slide_example.pptx"
    }
    
    print("\nRequest Payload:")
    print(json.dumps(payload, indent=2))
    
    try:
        response = requests.post(url, json=payload)
        response.raise_for_status()
        
        print("\nResponse:")
        print(json.dumps(response.json(), indent=2))
        print("\n✅ Points slide generated successfully!")
        
    except requests.exceptions.RequestException as e:
        print(f"\n❌ Error: {e}")
        if hasattr(e.response, 'text'):
            print(f"Details: {e.response.text}")


def example_2_image_text_slide():
    """Example: Generate an image+text slide"""
    print("\n" + "="*60)
    print("EXAMPLE 2: Generate Image+Text Slide")
    print("="*60)
    
    url = f"{BASE_URL}/generate-image-text-slide"
    
    payload = {
        "template_s3_url": "presentations/template.pptx",
        "slide_data": {
            "slide_number": 3,
            "title": "System Architecture",
            "text": "Our microservices architecture enables scalable and resilient API integration. The system uses event-driven communication patterns and implements circuit breakers for fault tolerance.",
            "image_url": "https://via.placeholder.com/800x600/4A90E2/FFFFFF?text=Architecture+Diagram",
            "title_color": "#3667B2",
            "text_color": "#000000"
        },
        "upload_to_s3": True,
        "output_filename": "image_text_slide_example.pptx"
    }
    
    print("\nRequest Payload:")
    print(json.dumps(payload, indent=2))
    
    try:
        response = requests.post(url, json=payload)
        response.raise_for_status()
        
        print("\nResponse:")
        print(json.dumps(response.json(), indent=2))
        print("\n✅ Image+Text slide generated successfully!")
        
    except requests.exceptions.RequestException as e:
        print(f"\n❌ Error: {e}")
        if hasattr(e.response, 'text'):
            print(f"Details: {e.response.text}")


def example_3_table_slide():
    """Example: Generate a table slide"""
    print("\n" + "="*60)
    print("EXAMPLE 3: Generate Table Slide")
    print("="*60)
    
    url = f"{BASE_URL}/generate-table-slide"
    
    payload = {
        "template_s3_url": "presentations/template.pptx",
        "slide_data": {
            "slide_number": 4,
            "title": "Cost Comparison Analysis",
            "table_data": [
                ["Feature", "Google API", "Third-Party Alternative", "Difference"],
                ["Monthly Base Cost", "$100", "$50", "-50%"],
                ["API Calls (per month)", "100,000", "50,000", "-50%"],
                ["Support Level", "Enterprise", "Standard", "-"],
                ["SLA Guarantee", "99.9%", "99.5%", "-0.4%"],
                ["Setup Time", "2 hours", "4 hours", "+2 hours"]
            ],
            "header_row": True,
            "header_color": "#3667B2"
        },
        "upload_to_s3": True,
        "output_filename": "table_slide_example.pptx"
    }
    
    print("\nRequest Payload:")
    print(json.dumps(payload, indent=2))
    
    try:
        response = requests.post(url, json=payload)
        response.raise_for_status()
        
        print("\nResponse:")
        print(json.dumps(response.json(), indent=2))
        print("\n✅ Table slide generated successfully!")
        
    except requests.exceptions.RequestException as e:
        print(f"\n❌ Error: {e}")
        if hasattr(e.response, 'text'):
            print(f"Details: {e.response.text}")


def example_4_multi_slide():
    """Example: Generate a complete presentation with multiple slide types"""
    print("\n" + "="*60)
    print("EXAMPLE 4: Generate Multi-Slide Presentation")
    print("="*60)
    
    url = f"{BASE_URL}/generate-multi-slide"
    
    payload = {
        "template_s3_url": "presentations/template.pptx",
        "slides_config": [
            {
                "slide_type": "points",
                "slide_data": {
                    "slide_number": 2,
                    "header": "Executive Summary",
                    "description": "Key findings from our API integration analysis",
                    "image_url": "https://via.placeholder.com/600x400/3667B2/FFFFFF?text=Executive+Summary",
                    "points": [
                        {"text": "Google API provides better reliability"},
                        {"text": "Third-party option is more cost-effective"},
                        {"text": "Both meet performance requirements"}
                    ],
                    "header_color": "#3667B2"
                }
            },
            {
                "slide_type": "image_text",
                "slide_data": {
                    "slide_number": 3,
                    "title": "Integration Architecture",
                    "text": "The proposed architecture leverages cloud-native patterns for maximum scalability and reliability.",
                    "image_url": "https://via.placeholder.com/800x600/4A90E2/FFFFFF?text=Architecture",
                    "title_color": "#3667B2"
                }
            },
            {
                "slide_type": "table",
                "slide_data": {
                    "slide_number": 4,
                    "title": "Detailed Cost Analysis",
                    "table_data": [
                        ["Item", "Q1", "Q2", "Q3", "Q4"],
                        ["API Costs", "$100", "$120", "$150", "$180"],
                        ["Infrastructure", "$50", "$50", "$60", "$70"],
                        ["Support", "$30", "$30", "$30", "$30"],
                        ["Total", "$180", "$200", "$240", "$280"]
                    ],
                    "header_row": True,
                    "header_color": "#3667B2"
                }
            }
        ],
        "upload_to_s3": True,
        "output_filename": "complete_presentation_example.pptx"
    }
    
    print("\nRequest Payload:")
    print(json.dumps(payload, indent=2))
    
    try:
        response = requests.post(url, json=payload)
        response.raise_for_status()
        
        print("\nResponse:")
        print(json.dumps(response.json(), indent=2))
        print("\n✅ Multi-slide presentation generated successfully!")
        
    except requests.exceptions.RequestException as e:
        print(f"\n❌ Error: {e}")
        if hasattr(e.response, 'text'):
            print(f"Details: {e.response.text}")


def example_5_get_slide_types():
    """Example: Get supported slide types"""
    print("\n" + "="*60)
    print("EXAMPLE 5: Get Supported Slide Types")
    print("="*60)
    
    url = f"{BASE_URL}/slide-types"
    
    try:
        response = requests.get(url)
        response.raise_for_status()
        
        print("\nResponse:")
        print(json.dumps(response.json(), indent=2))
        
    except requests.exceptions.RequestException as e:
        print(f"\n❌ Error: {e}")
        if hasattr(e.response, 'text'):
            print(f"Details: {e.response.text}")


def main():
    """Run all examples"""
    print("\n" + "#"*60)
    print("# Slide Generation API - Examples")
    print("#"*60)
    print("\nMake sure the API server is running at", BASE_URL)
    print("Start with: python main.py")
    
    # Check if API is available
    try:
        health_check = requests.get("http://localhost:8000/health")
        if health_check.status_code == 200:
            print("✅ API server is running!")
        else:
            print("⚠️  API server returned unexpected status")
            return
    except requests.exceptions.RequestException:
        print("❌ Cannot connect to API server. Please start it first.")
        return
    
    # Run examples
    # Uncomment the examples you want to run
    
    # Example 1: Points slide
    # example_1_points_slide()
    
    # Example 2: Image+Text slide
    # example_2_image_text_slide()
    
    # Example 3: Table slide
    # example_3_table_slide()
    
    # Example 4: Multi-slide presentation
    # example_4_multi_slide()
    
    # Example 5: Get slide types info
    example_5_get_slide_types()
    
    print("\n" + "="*60)
    print("Examples completed!")
    print("="*60)
    print("\nTips:")
    print("1. Uncomment the examples you want to run in the main() function")
    print("2. Replace 'presentations/template.pptx' with your actual S3 template path")
    print("3. Replace placeholder image URLs with your actual image URLs")
    print("4. Check the generated files in S3 or as downloads")
    print("\n")


if __name__ == "__main__":
    main()
