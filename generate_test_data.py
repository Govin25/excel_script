import pandas as pd
import random

def generate_questions(num_questions=1000):
    # List of question templates and topics
    templates = [
        "What is {}?",
        "How does {} work?",
        "Why is {} important?",
        "Can you explain {}?",
        "What are the benefits of {}?",
        "How to {}?",
        "What are the main components of {}?",
        "What is the difference between {} and {}?",
        "What are the best practices for {}?",
        "How has {} evolved over time?"
    ]
    
    topics = [
        "artificial intelligence", "machine learning", "data science", "cloud computing",
        "cybersecurity", "blockchain", "internet of things", "5G technology",
        "quantum computing", "virtual reality", "augmented reality", "robotics",
        "sustainable energy", "climate change", "renewable resources", "space exploration",
        "genetic engineering", "biotechnology", "nanotechnology", "3D printing",
        "digital marketing", "social media", "content creation", "e-commerce",
        "project management", "leadership", "team building", "communication skills",
        "time management", "productivity", "work-life balance", "stress management",
        "personal finance", "investment strategies", "stock market", "cryptocurrency",
        "healthy eating", "exercise", "mental health", "sleep hygiene",
        "programming", "software development", "web design", "mobile apps",
        "photography", "video editing", "graphic design", "music production",
        "languages", "cultural diversity", "globalization", "international relations",
        "education", "online learning", "skill development", "career growth",
        "sports", "fitness", "nutrition", "wellness",
        "travel", "tourism", "adventure", "exploration"
    ]
    
    questions = []
    for _ in range(num_questions):
        template = random.choice(templates)
        if "{}" in template:
            if template.count("{}") == 2:
                # For templates with two placeholders
                topic1 = random.choice(topics)
                topic2 = random.choice([t for t in topics if t != topic1])
                question = template.format(topic1, topic2)
            else:
                # For templates with one placeholder
                question = template.format(random.choice(topics))
        else:
            question = template.format(random.choice(topics))
        questions.append(question)
    
    return questions


def main():
    # Generate 50 questions
    num_questions = 50
    questions = generate_questions(num_questions)
    
    # Create DataFrame
    df = pd.DataFrame(questions, columns=['Questions'])
    
    # Save to Excel
    output_file = f'test_questions_{num_questions}.xlsx'
    df.to_excel(output_file, index=False)
    print(f"Generated {len(questions)} questions and saved to {output_file}")

if __name__ == "__main__":
    main() 