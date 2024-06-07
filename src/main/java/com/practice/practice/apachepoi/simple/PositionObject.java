package com.practice.practice.apachepoi.simple;

import lombok.EqualsAndHashCode;
import lombok.Getter;
import lombok.Setter;
import lombok.ToString;

/**
 * CellController의 Image File 또는 Object File의 위치를 조정하기 위해 사용.
 */
@Getter
@Setter
@EqualsAndHashCode
@ToString
public class PositionObject {
    private int dx1;
    private int dy1;
    private int dx2;
    private int dy2;

    public PositionObject(int dx1, int dy1, int dx2, int dy2) {
        if(dx1 < 0 || dy1 < 0 || dx2 < 0 || dy2 < 0){
            throw new IllegalArgumentException("Values of the parameter must be greater than or equal to zero.");
        }
        this.dx1 = dx1;
        this.dy1 = dy1;
        this.dx2 = dx2;
        this.dy2 = dy2;
    }
}
