"""
Test DOCX Generator
"""

import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from generators.docx_generator import DocxGenerator
from agents.writing_agent import WritingAgent

def test_simple_document():
    """Test creating a simple document"""
    
    print("=" * 70)
    print("TEST 1: SIMPLE DOCUMENT")
    print("=" * 70)
    print()
    
    generator = DocxGenerator()
    
    # Create simple document
    filepath = generator.create_simple_document(
        title="Test Document",
        content="This is a test document created by DocFlow AI. It demonstrates the basic document generation capability."
    )
    
    print(f"\n✅ Simple document created: {filepath}")
    print()

def test_ai_generated_document():
    """Test creating document with AI-generated content"""
    
    print("=" * 70)
    print("TEST 2: AI-GENERATED DOCUMENT")
    print("=" * 70)
    print()
    
    # Generate content with AI
    print("Step 1: Generating content with WritingAgent...")
    agent = WritingAgent()
    
    task = {
        'type': 'executive_summary',
        'topic': 'Q4 2024 Sales Performance',
        'data': {
            'Total Revenue': '$2.5M',
            'Growth': '23% YoY',
            'Top Region': 'North America',
            'Key Achievement': 'Closed 3 enterprise deals'
        },
        'tone': 'professional',
        'length': 'medium'
    }
    
    result = agent.process(task)
    
    if result['status'] == 'success':
        print(f"✓ Content generated ({result['word_count']} words)")
        print()
        
        # Create document with AI content
        print("Step 2: Creating Word document...")
        generator = DocxGenerator()
        
        doc_content = {
            'title': 'Q4 2024 Sales Performance',
            'subtitle': 'Executive Summary',
            'sections': [
                {
                    'heading': 'Overview',
                    'content': result['content'],
                    'type': 'paragraph'
                }
            ],
            'metadata': {
                'author': 'DocFlow AI',
                'company': 'Your Company',
                'date': 'December 2024'
            }
        }
        
        filepath = generator.create_document(doc_content)
        
        print()
        print("=" * 70)
        print("✅ SUCCESS!")
        print("=" * 70)
        print(f"\n📄 Document created: {filepath}")
        print("\nYou can now open this Word document!")
        print()
    else:
        print(f"❌ AI generation failed: {result.get('error')}")

def test_multi_section_document():
    """Test creating document with multiple sections"""
    
    print("=" * 70)
    print("TEST 3: MULTI-SECTION DOCUMENT WITH AI")
    print("=" * 70)
    print()
    
    agent = WritingAgent()
    generator = DocxGenerator()
    
    # Generate multiple sections
    sections_data = [
        {
            'heading': 'Executive Summary',
            'task': {
                'type': 'executive_summary',
                'topic': 'Q4 Performance',
                'data': {'Revenue': '$2.5M', 'Growth': '23%'},
                'tone': 'professional',
                'length': 'short'
            }
        },
        {
            'heading': 'Key Achievements',
            'task': {
                'type': 'detailed_section',
                'topic': 'Major Wins in Q4',
                'data': {
                    'Enterprise Deals': '3 closed',
                    'New Customers': '45',
                    'Retention Rate': '94%'
                },
                'tone': 'professional',
                'length': 'short'
            }
        },
        {
            'heading': 'Looking Forward',
            'task': {
                'type': 'detailed_section',
                'topic': 'Q1 2025 Objectives',
                'data': {
                    'Target Revenue': '$3M',
                    'New Hires': '5 sales reps',
                    'Market Expansion': 'EMEA region'
                },
                'tone': 'professional',
                'length': 'short'
            }
        }
    ]
    
    # Generate all sections
    print("Generating content for all sections...")
    sections = []
    
    for i, section_data in enumerate(sections_data):
        print(f"\n  Section {i+1}/{len(sections_data)}: {section_data['heading']}")
        result = agent.process(section_data['task'])
        
        if result['status'] == 'success':
            sections.append({
                'heading': section_data['heading'],
                'content': result['content'],
                'type': 'paragraph'
            })
            print(f"  ✓ Generated ({result['word_count']} words)")
        else:
            print(f"  ✗ Failed")
    
    # Create document
    print("\nCreating multi-section document...")
    
    doc_content = {
        'title': 'Q4 2024 Business Report',
        'subtitle': 'Quarterly Performance Review',
        'sections': sections,
        'metadata': {
            'author': 'DocFlow AI',
            'company': 'Your Company',
            'date': 'December 2024'
        }
    }
    
    filepath = generator.create_document(doc_content)
    
    print()
    print("=" * 70)
    print("✅ SUCCESS!")
    print("=" * 70)
    print(f"\n📄 Multi-section document created: {filepath}")
    print(f"   Sections: {len(sections)}")
    print("\nYou can now open this professional report!")
    print()

if __name__ == "__main__":
    # Run all tests
    test_simple_document()
    print("\n" + "="*70 + "\n")
    
    test_ai_generated_document()
    print("\n" + "="*70 + "\n")
    
    test_multi_section_document()