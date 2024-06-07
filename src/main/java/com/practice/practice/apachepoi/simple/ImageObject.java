package com.practice.practice.apachepoi.simple;

import lombok.Getter;
import java.awt.Image;
import javax.imageio.ImageIO;
import java.io.ByteArrayInputStream;
import java.io.IOException;

@Getter
public class ImageObject {
    private byte[] imageByteArray;
    private ImageFormatEnum imageFormatEnum;
    private String imageKey;
    private Image image;
    
    public ImageObject(byte[] imageByteArray, ImageFormatEnum imageFormatEnum, String imageKey) throws IOException{
        this.imageByteArray = imageByteArray;
        this.imageFormatEnum = imageFormatEnum;
        this.imageKey = imageKey;
        this.image = ImageIO.read(new ByteArrayInputStream(imageByteArray));
    }

    public int getWidth(){
        return image.getWidth(null);
    }

    public int getHeight(){
        return image.getHeight(null);
    }
}