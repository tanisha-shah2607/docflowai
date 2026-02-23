"""
Base Agent Class - Using Groq (Free & Fast!)
"""

from groq import Groq
from typing import Dict, Any
import sys
import os

# Add parent directory to path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from config import config

class BaseAgent:
    """
    Base class for all agents in DocFlow AI
    Uses Groq API (free and fast)
    """
    
    def __init__(self, name: str, role: str):
        """
        Initialize the agent
        
        Args:
            name: Agent's name (e.g., "WritingAgent")
            role: What this agent does
        """
        self.name = name
        self.role = role
        
        # Validate config
        config.validate()
        
        # Initialize Groq client
        self.client = Groq(api_key=config.GROQ_API_KEY)
        self.model = config.GROQ_MODEL
        
        print(f"✓ {self.name} initialized - {self.role}")
    
    def call_ai(self, prompt: str, max_tokens: int = 4000) -> str:
        """
        Call Groq AI with a prompt
        
        Args:
            prompt: What to ask the AI
            max_tokens: Maximum length of response
            
        Returns:
            AI's response as text
        """
        try:
            chat_completion = self.client.chat.completions.create(
                messages=[
                    {
                        "role": "user",
                        "content": prompt,
                    }
                ],
                model=self.model,
                temperature=0.7,
                max_tokens=max_tokens,
                top_p=1,
                stream=False,
            )
            
            return chat_completion.choices[0].message.content
            
        except Exception as e:
            print(f"❌ Error calling Groq: {str(e)}")
            raise
    
    def process(self, task: Dict[str, Any]) -> Dict[str, Any]:
        """
        Process a task - this will be overridden by specific agents
        
        Args:
            task: Dictionary containing task information
            
        Returns:
            Dictionary with results
        """
        raise NotImplementedError("Each agent must implement its own process method")