import React, { useMemo } from 'react';
import { Image } from 'react-konva';
import useImage from './hooks/useImage';
import { ShapeConfig } from 'konva/types/Shape';

export interface ImageProps extends ShapeConfig {
  url: string;
  spacing?: number
}

const ImageComponent: React.FC<ImageProps> = (props) => {
  let { url, width = 0, height = 0, x = 0, y = 0, spacing = 1, ...rest } = props
  const { image, width: imageWidth, height: imageHeight, status } = useImage({ url });  
  const aspectRatio = useMemo(() => {
    return Math.min((width - spacing) / imageWidth, (height - spacing) / imageHeight)
  }, [ imageWidth, imageHeight, width, height, spacing])
  x = x + spacing
  y = y + spacing
  
  width = (aspectRatio * imageWidth)
  height = (aspectRatio * imageHeight)
  if (status !== 'loaded') {
    return null
  }
  return (
    <Image {...rest} x={x} y={y} height={height} width={width} image={image} />
  );
}

export default ImageComponent
