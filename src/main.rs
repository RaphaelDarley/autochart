use calamine::{open_workbook, Error, RangeDeserializerBuilder, Reader, Xlsx};
use plotters::prelude::*;
use utils::*;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let config = ChartConfig {
        resolution: (640, 480),
        caption: Some("test caption"),
        source: Some("source"),
        watermark: Some("Raphael Darley"),
    };

    let book_path = r#"A:\Raphael\projects\Graphs\bp-stats-review-2022-all-data.xlsx"#;
    let mut workbook: Xlsx<_> = open_workbook(book_path)?;

    let sheet_names: Vec<String> = workbook.sheet_names().to_owned();

    let range = workbook
        // .worksheet_range(&sheet_names[0])
        .worksheet_range("Gas Production - Bcm")
        .ok_or(Error::Msg("Can't find requested sheet"))??;

    // let mut iter = RangeDeserializerBuilder::new().from_range(&range)?;

    chart_series(
        "images/test_chart.png",
        vec![100.0, 101.0, 102.0],
        vec![1.0, 5.0, 7.0],
        config,
    )?;

    Ok(())
}

struct ChartConfig<'a> {
    resolution: (u32, u32),
    caption: Option<&'a str>,
    source: Option<&'a str>,
    watermark: Option<&'a str>,
}

fn chart_series(
    path: &str,
    x_axis: Vec<f32>,
    data: Vec<f32>,
    config: ChartConfig,
) -> Result<(), Box<dyn std::error::Error>> {
    let root = BitMapBackend::new(path, config.resolution).into_drawing_area();

    root.fill(&WHITE)?;

    let root = root.margin(10, 10, 10, 10);

    let caption = if let Some(caption) = config.caption {
        caption
    } else {
        ""
    };

    let (x_min, x_max) = min_max(&x_axis);

    let (y_min, y_max) = min_max(&data);

    let y_min = f32::min(y_min, 0f32);

    let mut chart = ChartBuilder::on(&root)
        .caption(caption, ("sans-serif", 40).into_font())
        .x_label_area_size(20)
        .y_label_area_size(40)
        // .build_cartesian_2d(x_min..x_max, y_min..y_max);
        .build_cartesian_2d(x_min..x_max, y_min..y_max)?;

    chart.configure_mesh().draw()?;

    let coords = x_axis.into_iter().zip(data.into_iter());

    chart.draw_series(LineSeries::new(coords, &RED))?;

    Ok(())
}

mod utils {
    pub fn min_max(vec: &[f32]) -> (f32, f32) {
        let mut min = f32::MAX;
        let mut max = f32::MIN;

        for n in vec {
            min = min.min(*n);
            max = max.max(*n);
        }
        (min, max)
    }
}
