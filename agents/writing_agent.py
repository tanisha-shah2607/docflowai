"""
Writing Agent - Generates document content using AI
"""

from agents.base_agent import BaseAgent
from typing import Dict, Any

class WritingAgent(BaseAgent):
    """
    Agent responsible for generating written content
    Uses Groq's Llama 3.3 model
    """
    
    def __init__(self):
        super().__init__(
            name="WritingAgent",
            role="Generates document content with appropriate tone and style"
        )
    
    def process(self, task: Dict[str, Any]) -> Dict[str, Any]:
        """
        Generate content based on task requirements
        
        Args:
            task: {
                'type': 'executive_summary' | 'detailed_section' | 'bullet_points',
                'topic': 'What to write about',
                'data': {key data points to include},
                'tone': 'professional' | 'casual' | 'technical',
                'length': 'short' | 'medium' | 'long'
            }
            
        Returns:
            {
                'content': 'Generated text',
                'word_count': int,
                'status': 'success' | 'error'
            }
        """
        
        print(f"\n📝 {self.name} processing: {task.get('type', 'unknown')}")
        
        try:
            # Build the prompt for AI
            prompt = self._build_prompt(task)
            
            # Call AI to generate content
            content = self.call_ai(prompt)
            
            # Calculate word count
            word_count = len(content.split())
            
            print(f"✓ Generated {word_count} words")
            
            return {
                'content': content,
                'word_count': word_count,
                'status': 'success'
            }
            
        except Exception as e:
            print(f"❌ Error in WritingAgent: {str(e)}")
            return {
                'content': '',
                'word_count': 0,
                'status': 'error',
                'error': str(e)
            }
    
    def _build_prompt(self, task: Dict[str, Any]) -> str:
        """Build a prompt for AI based on the task"""
        
        content_type = task.get('type', 'general')
        topic = task.get('topic', 'Unknown topic')
        data = task.get('data', {})
        tone = task.get('tone', 'professional')
        length = task.get('length', 'medium')
        
        # Length guidelines
        length_guide = {
            'short': '100-200 words',
            'medium': '300-500 words',
            'long': '800-1200 words'
        }
        
        # Build the prompt
        prompt = f"""You are a professional document writer for DocFlow AI.

Task: Write {content_type} content about "{topic}"

Tone: {tone}
Length: {length_guide.get(length, '300-500 words')}

Data to incorporate:
{self._format_data(data)}

Instructions:
1. Write clear, well-structured content
2. Use the data provided naturally in the narrative
3. Maintain {tone} tone throughout
4. Do not add headers unless specifically needed for structure
5. Focus on clarity, impact, and professional quality
6. Make it engaging and informative

Generate the content now:"""

        return prompt
    
    def _format_data(self, data: Dict[str, Any]) -> str:
        """Format data dictionary into readable text for the prompt"""
        if not data:
            return "No specific data provided"
        
        formatted = []
        for key, value in data.items():
            formatted.append(f"- {key}: {value}")
        
        return "\n".join(formatted)