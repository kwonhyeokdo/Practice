package com.practice.practice.apachepoi.simple;

import java.util.Map;

/**
 * 단위 변환 계산 클래스
 * 주로 static 필드/메서드 사용
 * 맑은 고딕 폰트를 기준으로 한다.
 */
class UnitConverter {
    public static String BASE_FONT_NAME = "맑은 고딕";
    public static double PIXEL_PER_POINT = 0.75;
    public static double POINT_PER_PIXEL = 1 / PIXEL_PER_POINT;
    public static int POI_WIDTH_UNIT = 256;
    public static int POI_HEIGHT_UNIT = 20;
    public static double MAX_HEIGHT_POINT = 409.5;
    public static double BASE_FONT_POINT = 10.0;
    public static Map<Integer, Double> CHARACTER_WIDTH_PIXEL_MAP = Map.ofEntries(
        Map.entry(5, 4.0),
        Map.entry(6, 4.0),
        Map.entry(7, 5.0),
        Map.entry(8, 6.0),
        Map.entry(9, 7.0),
        Map.entry(10, 7.0),
        Map.entry(11, 8.0),
        Map.entry(12, 9.0),
        Map.entry(13, 9.0),
        Map.entry(14, 10.0),
        Map.entry(15, 11.0),
        Map.entry(16, 12.0),
        Map.entry(17, 13.0),
        Map.entry(18, 13.0),
        Map.entry(19, 14.0),
        Map.entry(20, 15.0),
        Map.entry(21, 15.0)
        // Map.entry(22, 16.0),
        // Map.entry(23, 17.0),
        // Map.entry(24, 18.0)
    );
    public static Map<Integer, Integer> CHARACTER_HEIGHT_PIXEL_MAP = Map.ofEntries(
        Map.entry(5, 12),
        Map.entry(6, 13),
        Map.entry(7, 13),
        Map.entry(8, 16),
        Map.entry(9, 16),
        Map.entry(10, 18),
        Map.entry(11, 22),
        Map.entry(12, 23),
        Map.entry(13, 26),
        Map.entry(14, 27),
        Map.entry(15, 32),
        Map.entry(16, 35),
        Map.entry(17, 35),
        Map.entry(18, 35),
        Map.entry(19, 40),
        Map.entry(20, 42),
        Map.entry(21, 42)
        // Map.entry(22, 45),
        // Map.entry(23, 47),
        // Map.entry(24, 51)
    );

    /**
     * point를 pixel로 변환한다.
     * @param point
     * @return point를 pixel로 변환한다.
     */
    static int pointToPixel(double point){
        return (int)(Math.round(point * POINT_PER_PIXEL));
    }

    /**
     * Excel의 Column Width를 pixel로 변환한다.
     * @param width
     * @param fontPoint 넓이를 구할 Column의 기준 폰트 포인트
     * @return Excel의 Column Width를 pixel로 변환한다.
     */
    static int widthToWidthPixel(int width, double fontPoint){
        return (int)(width * CHARACTER_WIDTH_PIXEL_MAP.get((int)fontPoint) / POI_WIDTH_UNIT);
    }

    /**
     * pixel을 Excel의 width로 반환한다.
     * Excel의 width 특성상 완벽하게 pixel로 구현하기가 쉽지 않아서 오차가 있다.
     * @param widthPixel 변환할 넓이의 픽셀 값
     * @param fontPoint 넓이를 구할 Column의 기준 폰트 포인트
     * @return pixel을 Excel의 width로 반환한다.
     */
    static int widthPixelToWidth(int widthPixel, double fontPoint){
        return (int)Math.round(((double)widthPixel) / CHARACTER_WIDTH_PIXEL_MAP.get((int)fontPoint) * POI_WIDTH_UNIT);
    }

    /**
     * pixel을 Excel의 Height(point)로 반환한다.
     * @param point
     * @return pixel을 Excel의 Height(point)로 반환한다.
     */
    static short getHeightFromPoint(double point){
        // if(point > MAX_HEIGHT_POINT){
        //     point = MAX_HEIGHT_POINT;
        // }
        double height = point * POI_HEIGHT_UNIT;
        return (short)height;
    }

    /**
     * pixel을 Excel의 Height(point)로 반환한다.
     * @param pixel
     * @return pixel을 Excel의 width로 반환한다.
     */
    static short getHeightFromPixel(int pixel){
        double point = pixel * PIXEL_PER_POINT;
        return getHeightFromPoint(point);
    }
}
