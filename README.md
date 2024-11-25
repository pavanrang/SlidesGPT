# SlidesGPT

PowerPoint presentation generator built with Streamlit that creates professional presentations using AI content generation and optional image integration.


## Installation

1. Clone the repository:
```bash
git clone <repository-url>
cd ai-powerpoint-generator
```

2. Install required dependencies:
```bash
pip install -r requirements.txt
```

3. Set up your environment variables:
```bash
export GROQ_API_KEY="your_groq_api_key"
export SERPER_API_KEY="your_serper_api_key"
```

## Project Structure

```
.
├── GeneratedPresentations/    # Output directory for generated presentations
├── Designs/                   # PowerPoint template designs
│   ├── Design-1.pptx
│   ├── Design-2.pptx
│   └── ...
└── app.py                     # Main application file
```

## Usage

1. Start the Streamlit application:
```bash
streamlit run app.py
```

2. Access the web interface at `http://localhost:8501`

3. Enter your presentation topic and customize settings:
   - Specify the presentation topic
   - Choose a design template (1-7)
   - Toggle image inclusion
   - Click "Generate Presentation"

4. Download the generated PowerPoint file

## Features in Detail

### Content Generation
- Uses Groq's Mixtral-8x7b-32768 model for high-quality content
- Structured content with titles, subtitles, and bullet points
- Hierarchical formatting with main points and sub-points
- Automatic slide organization and flow

### Design Templates
- 7 professional design templates
- Consistent formatting and styling
- Support for title slides and content slides
- Dynamic layout selection for visual variety

### Image Integration
- Google Image Search integration via Serper API
- Automatic image placement and sizing
- Relevant images based on slide content
- Optional toggle for image inclusion

### Text Formatting
- Hierarchical bullet points
- Consistent font sizes and spacing
- Professional slide layouts
- Proper text alignment and positioning

## API Integration

### Groq API
Used for content generation with the following features:
- Model: Mixtral-8x7b-32768
- Temperature: 0.7
- Max tokens: 4000

### Google Serper API
Used for image search with the following features:
- Image search endpoint
- Returns top 5 relevant images
- Automatic error handling and fallback



## Limitations

- The mixtral model that i use may not generate structred output everytime, since the guardrails are setup only through prompt engineering
