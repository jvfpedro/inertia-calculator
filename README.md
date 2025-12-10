# inertia-calculator

This project is a Python-based tool designed to calculate the moment of inertia of various sections in PNG format. This project is developed as part of my studies in the "Computational Tools for Civil Engineering" course at the Federal University of Santa Catarina (UFSC), Brazil.

## Description

The tool processes images of sections and calculates the moment of inertia, as well as the center of gravity (CG) position. It supports PNG images and provides the following:

- Calculates the moment of inertia (I) and the position of the center of gravity (Y).
- Visualizes the section with the CG line overlaid.
- Generates a Word document with the calculated results and images.

The project uses several Python libraries, including **customtkinter**, **PIL**, **matplotlib**, and **python-docx**.

## Features

- **Image Processing**: Reads PNG images and processes them to calculate geometric properties.
- **Calculation of Moment of Inertia**: Calculates the moment of inertia for the given section.
- **Center of Gravity**: Computes the height of the center of gravity from the base.
- **Visualization**: Displays the section and overlays the center of gravity line.
- **Word Report**: Generates a DOCX file with the results, images, and other relevant information.

## Libraries Used

- `numpy`
- `PIL` (Pillow)
- `matplotlib`
- `python-docx`
- `customtkinter`

## Installation

1. Clone this repository to your local machine:
   ```bash
   git clone https://github.com/danieltanjos/InertiaCalc.git

2. Install the required dependencies:
   ```bash
   pip install numpy pillow matplotlib python-docx customtkinter

## Usage

1. Run the application:
   ```bash
   python GUI.py
2. Select a PNG image of the section.
3. Enter the maximum vertical dimension of the section (in mm).
4. Press the "Calculate" button to compute the moment of inertia and center of gravity.
5. Optionally, generate a DOCX report with the results by clicking the "Generate DOCX" button.

## Example Output

After processing the image, the tool will display:
- The width and height of the image.
- The calculated center of gravity (CG) height.
- The calculated moment of inertia.
  
Additionally, a report is generated in Word format with:
- The image of the section.
- The calculated values.
- A second image with the CG line overlaid.

Inside the "dados" folder, there are example PNG files that can be used to test the project.

## Acknowledgments
This project was developed as part of my coursework for the "Ferramentas Computacionacias para Engenharia Civil" (Computational Tools for Civil Engineering) course at **UFSC** (Federal University of Santa Catarina).

## Author

João Vitor Ferreira Pedro
[Civil Engineer - UFSC]
[https://github.com/jvfpedro] [jvfpedro@gmail.com]

Daniel Tavares dos Anjos
[collaborator]
[https://github.com/danieltanjos]
