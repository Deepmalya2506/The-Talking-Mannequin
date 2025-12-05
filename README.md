# The Talking Mannequin — Neural Twin

**The future of interaction is watching you.**

![Demo Screenshot](./Screenshot%202025-12-05%20121900.png)

## Overview

Neural Twin (The Talking Mannequin) is a deskside, human-like AI companion that understands words and emotions, presents information naturally, and responds in real time. It is designed for both development machines (laptop / Windows) and edge devices (Raspberry Pi 4B running Debian/Raspberry Pi OS). The project focuses on agentic presentation assistance: the mannequin can present slides, answer context-aware questions during a talk, and adapt behavior based on detected emotion and gestures.

> Status: In progress — Dockerization and Raspberry Pi 4B integration are actively being worked on. Contributions are highly solicited.

---

## Table of Contents

- About
- Key Features
- Repository Structure
- Quickstart (Laptop)
- Quickstart (Raspberry Pi)
- Installation
- Running
- Configuration
- Architecture & Components
- Usage Examples
- Requirements
- Development & Contributing
- Roadmap
- Troubleshooting
- Security & Privacy
- License
- Acknowledgements & Contact

---

## About

Neural Twin is more than a chatbot. It is an embodied, multimodal assistant that listens, understands, and interacts like a real partner. Primary goals:

- Agentic presentation: assist presenters by controlling slides, providing context-aware explanations, and fielding audience questions.
- Multimodal understanding: combine speech, vision (optional), and sentiment detection to adapt tone and expressions.
- Offline-first: run locally on devices such as Raspberry Pi 4B for privacy and availability.
- Extensible: modular design to swap models, add skills, or integrate new hardware.

## Key Features

- Natural language understanding for conversational interaction
- Speech-to-text and text-to-speech pipelines (pyttsx / pyttsx3 recommended)
- Emotion and sentiment detection for adaptive responses
- Presentation control (next/previous slide, explain slide, highlight sections)
- Modular hardware interface for servos, LEDs, cameras, and microphones
- Separate environments and requirements for laptop and Raspberry Pi

## Repository Structure

- README.md (this file)
- LICENSE
- main_pi.py           — Primary entry point for Raspberry Pi deployments
- laptop_client.py     — Lightweight client entrypoint for laptop development
- debadree.py          — Pi/DEB-specific utilities and scripts
- hw.py                — Hardware abstraction layer and device control
- testing.py           — Testing and experimentation scripts
- requirements_laptop.txt  — pip requirements for laptop / Windows environments
- requirements_pi.txt      — pip requirements tailored for Raspberry Pi (Debian)
- documents/           — Place presentations and documents here
- documents.db         — Local SQLite DB used for storing documents / embeddings
- Screenshot 2025-12-05 121900.png  — Visual demo


## Quickstart (Laptop)

1. Clone the repository:
   ```bash
   git clone https://github.com/Deepmalya2506/The-Talking-Mannequin.git
   cd The-Talking-Mannequin
   ```

2. Create and activate a virtual environment (recommended):
   ```bash
   python -m venv venv
   # Windows
   venv\Scripts\activate
   # macOS / Linux
   source venv/bin/activate
   ```

3. Install laptop dependencies:
   ```bash
   pip install -r requirements_laptop.txt
   ```

4. Run the laptop client (development / testing):
   ```bash
   python laptop_client.py
   ```


## Quickstart (Raspberry Pi 4B)

These instructions assume Raspberry Pi OS (Debian-based). The repository includes a Raspberry Pi-specific entrypoint main_pi.py and a requirements_pi.txt tuned for the Pi.

1. Update and upgrade your Pi:
   ```bash
   sudo apt update && sudo apt upgrade -y
   ```

2. Install system dependencies commonly required for audio/video and hardware access:
   ```bash
   sudo apt install -y python3-venv python3-pip build-essential libatlas-base-dev libasound2-dev libportaudio2 ffmpeg
   ```

3. Clone and prepare the project:
   ```bash
   git clone https://github.com/Deepmalya2506/The-Talking-Mannequin.git
   cd The-Talking-Mannequin
   python3 -m venv venv
   source venv/bin/activate
   pip install -r requirements_pi.txt
   ```

4. Run the Pi entrypoint (the program will attempt to use attached microphone/camera and available hardware):
   ```bash
   python3 main_pi.py
   ```


## Installation (General)

- The repository contains two requirements files: requirements_laptop.txt and requirements_pi.txt. Use the appropriate one for your environment.
- For on-device models consider quantized/edge-optimized variants or hardware accelerators (Coral, NPU, etc.).

## Running

- Laptop: python laptop_client.py
- Raspberry Pi: python3 main_pi.py

If you encounter permission errors for audio/video devices, ensure your user has device access and that ALSA / PulseAudio are configured correctly.

## Configuration

Create a config/ directory (optional) to place YAML or JSON settings for:

- audio devices (mic, speakers)
- hardware pins and servo mappings
- default models and model paths
- logging verbosity

Use environment variables to toggle modes (e.g., EDGE_MODE=pi or DEV_MODE=laptop).


## Architecture & Components

Conceptual layers:

- Input: microphone, camera, and other sensors
- Perception: ASR (speech-to-text), vision/gesture detectors, sentiment analysis
- Understanding: intent detection, context management, conversation state
- Action: TTS (pyttsx / pyttsx3), presentation control, hardware actuation (servos/LEDs)
- Storage: local SQLite (documents.db) for documents, session information, and embeddings

Primary modules:

- hw.py: hardware abstraction and low-level device control
- main_pi.py: orchestration and Pi-specific runtime
- laptop_client.py: lightweight entrypoint for development and demos
- debadree.py & testing.py: utility and experimental code


## Usage Examples

- Presentation mode:
  - Run the Pi with presentation files in documents/, then ask aloud:
    - "Next slide" — mannequin advances and provides a short summary
    - "Explain this chart" — attempts to parse slide content and summarize

- Conversational mode:
  - "Who is the presenter?" — reads local metadata or session memory
  - "Summarize the last slide" — agent uses the document store (documents.db) to summarize content


## Requirements

- Python 3.9+
- See requirements_laptop.txt and requirements_pi.txt for environment-specific packages

Hardware suggestions:
- Laptop: USB microphone, webcam (optional)
- Pi: Raspberry Pi 4B (4GB+ recommended), USB mic / Pi-compatible microphone, camera module, stable power supply


## Development & Contributing

This project is actively under development. Contributions are welcome — especially for:

- Raspberry Pi performance and packaging
- Dockerization and deployment scripts
- Plug-ins for new TTS/ASR backends
- Gesture and facial expression modules

How to contribute:
1. Fork the repository
2. Create a feature branch: git checkout -b feat/your-feature
3. Add tests and documentation for your change
4. Open a pull request describing the change and motivation

Please include a short description of how you tested changes on either a laptop or Raspberry Pi.

> NOTE: The project is still in progress. Any contributions, bug reports, or suggestions are highly solicited and appreciated.


## Roadmap

### Near-term
- Complete Dockerization
- Improve runtime stability on Raspberry Pi 4B
- Integrate quantized offline models for on-device performance

### Medium-term
- Gesture and facial expression detection and richer non-verbal responses
- Session memory and personalization modules
- Plugin system for custom skills and connectors

### Long-term
- Robust agentic presentation capabilities across multiple platforms
- Optional user-consented cloud-sync for cross-device personalization


## Troubleshooting

- Audio device errors: ensure microphone is present and accessible. On Raspberry Pi, check ALSA and user groups (audio).
- Module import issues: confirm virtual environment is activated and correct requirements file was used.
- High CPU usage / out-of-memory: reduce active modules (disable vision), use smaller models, or increase swap cautiously on the Pi.


## Security & Privacy

- Offline-first design minimizes data exposure.
- If cloud integrations are added, document what data leaves the device and obtain user consent.


## Credits & Third-party Libraries

- TTS: pyttsx / pyttsx3 (recommended)
- Other libraries: see requirements files for details


## Contact

Maintained by Deepmalya2506. Open issues, feature requests, and PRs are welcome.


---
