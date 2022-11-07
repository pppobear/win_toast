pub mod notification;

#[cfg(test)]
mod tests {
    use std::path::Path;
    use crate::notification::*;

    #[test]
    fn test_struct_toast() -> windows::core::Result<()> {
        let toast = Toast {
            app_id: Toast::POWERSHELL_APP_ID.to_string(),
            binding_elems: vec![BindingElem {
                hero: Some(Image {
                    alt: "logo".to_string(),
                    src: Path::new(env!("CARGO_MANIFEST_DIR")).join("resources/test/flower.jpeg").display().to_string(),
                    hint_crop: HintCrop::Circle,
                    hint_align: ImageAlign::Default,
                }),
                icon: Some(Image {
                    alt: "logo".to_string(),
                    src: Path::new(env!("CARGO_MANIFEST_DIR")).join("resources/test/chick.jpeg").display().to_string(),
                    hint_crop: HintCrop::Circle,
                    hint_align: ImageAlign::Default,
                }),
                title: ("hello".to_string(), Some(1)),
                elems: vec![BindingInnerElement::Text("Hello, win_toast.".into())],
            }],
            duration: Duration::Short,
            scenario: Scenario::Default,
            audio: None,
            actions: vec![],
            use_btn_style: false,
            display_timestamp: None,
        };
        toast.show()?;
        std::thread::sleep(std::time::Duration::from_millis(10));
        Ok(())
    }

    #[test]
    fn test_text_xml_toast() -> windows::core::Result<()> {
        Toast::show_with_xml(Toast::POWERSHELL_APP_ID, r#"
        "#)?;
        std::thread::sleep(std::time::Duration::from_millis(10));
        Ok(())
    }
}
