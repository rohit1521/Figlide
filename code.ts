// This plugin will open a window to prompt the user to enter a number, and
// it will then create that many rectangles on the screen.

// This file holds the main code for plugins. Code in this file has access to
// the *figma document* via the figma global object.
// You can access browser APIs in the <script> tag inside "ui.html" which has a
// full browser environment (See https://www.figma.com/plugin-docs/how-plugins-run).

// This shows the HTML page in "ui.html".
figma.showUI(__html__, { width: 1000, height: 600 });

// Add interfaces for type safety
interface FrameData {
  name: string;
  width: number;
  height: number;
  textCount: number;
  imageCount: number;
  aspectRatioWarning: boolean;
  thumbnail: string;
  selectionIndex: number;
  backgroundColor?: RGB;
}

interface FrameRequest {
  id: string;
  selectionIndex: number;
}

interface FrameElement {
  type: 'TEXT' | 'IMAGE';
  x: number;
  y: number;
  width: number;
  height: number;
}

interface TextElement extends FrameElement {
  type: 'TEXT';
  text: string;
  fontSize: number;
  fontName?: string;
  fontColor?: string;
  fontWeight: number;
}

interface ImageElement extends FrameElement {
  type: 'IMAGE';
  imageBytes: Uint8Array;
}

// Function to gather frame data
async function getFrameData(frame: FrameNode, index: number): Promise<FrameData> {
  const imageBytes = await frame.exportAsync({ 
    format: "PNG", 
    constraint: { type: "SCALE", value: 0.25 } 
  });
  
  const aspectRatio = frame.width / frame.height;
  
  // Safer type checking for fills
  const background = frame.backgrounds[0];
  const backgroundColor = background && background.type === 'SOLID' 
    ? background.color 
    : undefined;

  return {
    name: frame.name,
    width: frame.width,
    height: frame.height,
    textCount: frame.findAll(node => node.type === 'TEXT').length,
    // Safer check for image fills
    imageCount: frame.findAll(node => {
      if (node.type === 'RECTANGLE') {
        const fills = node.fills as Paint[];
        return fills && fills.length > 0 && fills[0].type === 'IMAGE';
      }
      return false;
    }).length,
    aspectRatioWarning: Math.abs(aspectRatio - (16 / 9)) > 0.01,
    thumbnail: `data:image/png;base64,${figma.base64Encode(imageBytes)}`,
    selectionIndex: index,
    backgroundColor
  };
}

// Function to update UI with selected frames
async function updateSelectedFrames() {
  const selectedFrames = figma.currentPage.selection
    .filter(node => node.type === 'FRAME');
  
  try {
    const frameData = await Promise.all(
      selectedFrames.map((frame, index) => getFrameData(frame as FrameNode, index))
    );
    
    figma.ui.postMessage({ type: 'update-frames', frames: frameData });
  } catch (error) {
    console.error('Error updating frames:', error);
    figma.notify('Error updating frame selection');
  }
}

// Initial update and selection change listener
updateSelectedFrames();
figma.on('selectionchange', updateSelectedFrames);

// Process frame elements
async function processFrameElements(frame: FrameNode): Promise<(TextElement | ImageElement)[]> {
  const elements: (TextElement | ImageElement)[] = [];
  
  for (const node of frame.children) {
    try {
      if (node.type === 'TEXT') {
        const textElement = await processTextElement(node);
        if (textElement) elements.push(textElement);
      } 
      // Add handling for nested frames
      else if (node.type === 'FRAME') {
        // Export the frame as an image
        const imageBytes = await node.exportAsync({
          format: 'PNG',
          constraint: { type: 'SCALE', value: 1 }
        });
        
        elements.push({
          type: 'IMAGE',
          x: node.x,
          y: node.y,
          width: node.width,
          height: node.height,
          imageBytes
        });
      }
      else if (node.type === 'RECTANGLE') {
        const imageElement = await processImageElement(node);
        if (imageElement) elements.push(imageElement);
      }
      // You might want to add other node types here
    } catch (error) {
      console.error(`Error processing node in frame ${frame.name}:`, error);
    }
  }
  
  return elements;
}

// Process text elements
function processTextElement(node: TextNode): TextElement {
  const fills = node.fills as readonly Paint[];
  const fontName = node.fontName as FontName;
  
  return {
    type: 'TEXT',
    text: node.characters,
    x: node.x,
    y: node.y,
    width: node.width,
    height: node.height,
    fontSize: typeof node.fontSize === 'number' ? node.fontSize : 12,
    fontName: fontName?.family || 'Arial',
    fontColor: fills.length > 0 ? rgbToHex((fills[0] as SolidPaint).color) : undefined,
    fontWeight: typeof node.fontWeight === 'number' ? node.fontWeight : 400
  };
}

// Process image elements
async function processImageElement(node: RectangleNode): Promise<ImageElement | null> {
  const fills = node.fills as readonly Paint[];
  
  if (fills.length > 0 && fills[0].type === 'IMAGE') {
    const imageBytes = await node.exportAsync();
    return {
      type: 'IMAGE',
      x: node.x,
      y: node.y,
      width: node.width,
      height: node.height,
      imageBytes
    };
  }
  return null;
}

// Convert RGB color to hex
function rgbToHex(color: RGB): string {
  const r = Math.round(color.r * 255);
  const g = Math.round(color.g * 255);
  const b = Math.round(color.b * 255);
  return `#${((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1)}`;
}

// Handle messages from UI
figma.ui.onmessage = async (msg) => {
  console.log('Received message:', msg);
  
  switch (msg.type) {
    case 'create-thumbnail':
      await handleCreateThumbnail();
      break;
      
    case 'export-ppt':
      await handleExportPPT(msg.frames);
      break;
  }
};

// Handle thumbnail creation
async function handleCreateThumbnail() {
  const selection = figma.currentPage.selection;
  
  if (selection.length !== 1 || selection[0].type !== 'FRAME') {
    figma.notify('Please select a single frame.');
    return;
  }

  try {
    const frame = selection[0];
    const thumbnail = await frame.exportAsync({
      format: 'PNG',
      constraint: { type: 'SCALE', value: 1.2 },
      contentsOnly: true,
      useAbsoluteBounds: true,
    });
    
    figma.ui.postMessage({ 
      type: 'thumbnail', 
      data: figma.base64Encode(thumbnail),
      frameId: frame.id
    });
  } catch (error) {
    console.error('Export error:', error);
    figma.notify('Error creating thumbnail');
  }
}

// Handle PPT export
async function handleExportPPT(frames: FrameRequest[]) {
  try {
    const frameData = await Promise.all(frames.map(async (frame, index) => {
      const node = await figma.getNodeByIdAsync(frame.id);
      
      if (!node || node.type !== 'FRAME') {
        console.warn(`Invalid frame at index ${index}`);
        return null;
      }

      const elements = await processFrameElements(node);
      const background = node.backgrounds[0];
      const backgroundColor = background?.type === 'SOLID' 
        ? background.color 
        : undefined;

      return {
        selectionIndex: frame.selectionIndex,
        backgroundColor,
        elements
      };
    }));

    const validFrameData = frameData.filter(Boolean);
    
    if (validFrameData.length === 0) {
      throw new Error('No valid frames to process');
    }

    figma.ui.postMessage({ 
      type: 'frame-data',
      frameData: validFrameData
    });
  } catch (error) {
    console.error('Export error:', error);
    figma.notify('Export failed: ' + (error instanceof Error ? error.message : 'Unknown error'));
    figma.ui.postMessage({ 
      type: 'export-error',
      error: error instanceof Error ? error.message : 'Unknown error'
    });
  }
}
