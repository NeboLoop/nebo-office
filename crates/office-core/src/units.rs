/// Convert inches to DXA (twentieths of a point). 1 inch = 1440 DXA.
pub fn inches_to_dxa(inches: f64) -> i64 {
    (inches * 1440.0).round() as i64
}

/// Convert points to half-points (used for font sizes in OOXML). 1 pt = 2 half-points.
pub fn points_to_half_points(points: f64) -> i64 {
    (points * 2.0).round() as i64
}

/// Convert inches to EMUs (English Metric Units). 1 inch = 914400 EMU.
pub fn inches_to_emu(inches: f64) -> i64 {
    (inches * 914400.0).round() as i64
}

/// Convert points to twentieths of a point (used for spacing).
pub fn points_to_twips(points: f64) -> i64 {
    (points * 20.0).round() as i64
}

/// Convert DXA to inches.
pub fn dxa_to_inches(dxa: i64) -> f64 {
    dxa as f64 / 1440.0
}

/// Convert half-points to points.
pub fn half_points_to_points(hp: i64) -> f64 {
    hp as f64 / 2.0
}

/// Convert EMU to inches.
pub fn emu_to_inches(emu: i64) -> f64 {
    emu as f64 / 914400.0
}

/// Convert twips to points.
pub fn twips_to_points(twips: i64) -> f64 {
    twips as f64 / 20.0
}

/// Convert inches to points (used for PDF). 1 inch = 72 pt.
pub fn inches_to_points(inches: f64) -> f64 {
    inches * 72.0
}

/// Convert points to inches.
pub fn points_to_inches(points: f64) -> f64 {
    points / 72.0
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_inches_to_dxa() {
        assert_eq!(inches_to_dxa(1.0), 1440);
        assert_eq!(inches_to_dxa(0.5), 720);
        assert_eq!(inches_to_dxa(1.25), 1800);
    }

    #[test]
    fn test_points_to_half_points() {
        assert_eq!(points_to_half_points(12.0), 24);
        assert_eq!(points_to_half_points(10.5), 21);
    }

    #[test]
    fn test_inches_to_emu() {
        assert_eq!(inches_to_emu(1.0), 914400);
        assert_eq!(inches_to_emu(4.0), 3657600);
    }

    #[test]
    fn test_round_trip() {
        let inches = 1.25;
        assert!((dxa_to_inches(inches_to_dxa(inches)) - inches).abs() < 0.001);
        assert!((emu_to_inches(inches_to_emu(inches)) - inches).abs() < 0.001);
    }

    #[test]
    fn test_inches_to_points() {
        assert!((inches_to_points(1.0) - 72.0).abs() < 0.001);
        assert!((inches_to_points(0.5) - 36.0).abs() < 0.001);
        assert!((points_to_inches(inches_to_points(1.25)) - 1.25).abs() < 0.001);
    }
}
