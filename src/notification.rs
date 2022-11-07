use windows::core::*;
use windows::Data::Xml::Dom::{IXmlNode, XmlDocument, XmlElement};
use windows::UI::Notifications::{ToastNotification, ToastNotificationManager, ToastNotifier};
use crate::notification::HintCrop::Circle;

#[allow(dead_code)]
pub enum Duration {
    Default,
    /// 7 seconds
    Short,
    /// 25 seconds
    Long,
}

#[allow(dead_code)]
pub enum Scenario {
    /// The normal toast behavior.
    Default,
    /// This will be displayed pre-expanded and stay on the user's screen till dismissed. Audio will loop by default and will use alarm audio.
    Alarm,
    /// This will be displayed pre-expanded and stay on the user's screen till dismissed..
    Reminder,
    /// This will be displayed pre-expanded in a special call format and stay on the user's screen till dismissed. Audio will loop by default and will use ringtone audio.
    IncomingCall,
    /// An important notification. This allows users to have more control over what apps can send them high-priority toast notifications that can break through Focus Assist (Do not Disturb). This can be modified in the notifications settings.
    Urgent,
}

#[allow(dead_code)]
pub enum HintCrop {
    /// The image is not cropped and displayed as a square.
    Default,
    /// The image is cropped into a circle.
    Circle,
}

#[allow(dead_code)]
pub enum TextStyle {
    /// Default value. Style is determined by the renderer.
    Default,
    /// Smaller than paragraph font size.
    Caption,
    /// Same as Caption but with subtle opacity.
    CaptionSubtle,
    /// Paragraph font size.
    Body,
    /// Same as Body but with subtle opacity.
    BodySubtle,
    /// Paragraph font size, bold weight. Essentially the bold version of Body.
    Base,
    /// Same as Base but with subtle opacity.
    BaseSubtle,
    /// H4 font size.
    Subtitle,
    /// Same as Subtitle but with subtle opacity.
    SubtitleSubtle,
    /// H3 font size.
    Title,
    /// Same as Title but with subtle opacity.
    TitleSubtle,
    /// Same as Title but with top/bottom padding removed.
    TitleNumeral,
    /// H2 font size.
    Subheader,
    /// Same as Subheader but with subtle opacity.
    SubheaderSubtle,
    /// Same as Subheader but with top/bottom padding removed.
    SubheaderNumeral,
    /// H1 font size.
    Header,
    /// Same as Header but with subtle opacity.
    HeaderSubtle,
    /// Same as Header but with top/bottom padding removed.
    HeaderNumeral,
}

#[allow(dead_code)]
pub enum TextAlign {
    /// Default value. Alignment is automatically determined by the renderer.
    Default,
    /// Alignment determined by the current language and culture.
    Auto,
    /// Horizontally align the text to the left.
    Left,
    /// Horizontally align the text in the center.
    Center,
    /// Horizontally align the text to the right.
    Right,
}

#[allow(dead_code)]
pub enum ImageAlign {
    /// Default value. Alignment behavior determined by renderer.
    Default,
    /// Image stretches to fill available width (and potentially available height too, depending on where the image is placed).
    Stretch,
    /// Align the image to the left, displaying the image at its native resolution.
    Left,
    /// Align the image in the center horizontally, displayign the image at its native resolution.
    Center,
    /// Align the image to the right, displaying the image at its native resolution.
    Right,
}

#[allow(dead_code)]
pub enum SubgroupElement {
    Text_(InnerText),
    Image_(Image),
}

#[allow(dead_code)]
pub enum SoundSrc {
    /// Not support looping
    Default,
    /// Not support looping
    IM,
    /// Not support looping
    Mail,
    /// Not support looping
    Reminder,
    /// Not support looping
    SMS,
    Alarm,
    Alarm2,
    Alarm3,
    Alarm4,
    Alarm5,
    Alarm6,
    Alarm7,
    Alarm8,
    Alarm9,
    Alarm10,
    Call,
    Call2,
    Call3,
    Call4,
    Call5,
    Call6,
    Call7,
    Call8,
    Call9,
    Call10,
}

pub struct InnerText {
    pub text: String,
    pub hint_style: TextStyle,
    pub hint_warp: Option<bool>,
    pub hint_max_lines: Option<u32>,
    pub hint_min_lines: Option<u32>,
    pub hint_align: TextAlign,
}

pub struct Image {
    /// The URI of the image source, using one of these protocol handlers:
    ///
    /// - http:// or https://
    ///
    /// A web-based image.
    ///
    /// -  ms-appx:///
    ///
    /// An image included in the app package.
    ///
    /// - ms-appdata:///local/
    ///
    /// An image saved to local storage.
    ///
    /// - file:///
    ///
    /// A local image. (Supported only for desktop apps. This protocol cannot be used by UWP apps.)
    pub src: String,
    /// A description of the image, for users of assistive technologies.
    pub alt: String,
    /// The cropping of the image.
    ///
    /// - "circle" - The image is cropped into a circle.
    /// - Unspecified - The image is not cropped and displayed as a square.
    pub hint_crop: HintCrop,
    /// Only works for images inside a group/subgroup
    pub hint_align: ImageAlign,
}

#[allow(dead_code)]
pub enum BindingInnerElement {
    Text(String),
    Group(Vec<Vec<SubgroupElement>>),
    Image(Image),
}

#[allow(dead_code)]
pub enum InputType {
    /// The element is default text(optional).
    Text(Option<String>),
    /// The elements are (defaultSelectionBoxItemId and selections).
    ///
    /// The elements in the selection item are (id and title).
    Selection(Option<String>, Vec<(String, String)>),
}

pub struct Input {
    /// The ID associated with the input.
    pub id: String,
    /// The type of input.
    pub type_: InputType,
    /// The placeholder displayed for text input.
    pub place_holder_content: Option<String>,
    /// Text displayed as a label for the input.
    pub title: Option<String>,
}

#[allow(dead_code)]
pub enum ActivationType {
    /// Your foreground app is launched.
    Foreground,
    /// Your corresponding background task is triggered, and you can execute code in the background without interrupting the user.
    Background,
    /// Launch a different app using protocol activation.
    Protocol,
}

#[allow(dead_code)]
pub enum ActionPlacement {
    Default,
    ContextMenu,
}

pub struct Action {
    /// The content displayed on the button.
    pub content: String,
    /// App-defined string of arguments that the app will later receive if the user clicks this button.
    pub arguments: String,
    /// Decides the type of activation that will be used when the user interacts with a specific action.
    pub activation_type: ActivationType,
    /// When set to "contextMenu", the action becomes a context menu action added to
    /// the toast notification's context menu rather than a traditional toast button.
    pub placement: ActionPlacement,
    /// The URI of the image source for a toast button icon. These icons are white transparent 16x16 pixel images at 100% scaling and should have no padding included in the image itself. If you choose to provide icons on a toast notification, you must provide icons for ALL of your buttons in the notification, as it transforms the style of your buttons into icon buttons. Use one of the following protocol handlers:
    ///
    /// - http:// or https:// - A web-based image.
    ///
    /// - ms-appx:/// - An image included in the app package.
    ///
    /// - ms-appdata:///local/ - An image saved to local storage.
    ///
    /// - file:/// - A local image. (Supported only for desktop apps. This protocol cannot be used by UWP apps.)
    pub image_uri: Option<String>,
    /// Set to the Id of an input to position button beside the input.
    pub hint_input_id: Option<String>,
    /// The button style. useButtonStyle must be set to true in the toast element.
    ///
    /// - "Success" - The button is green
    ///
    /// - "Critical" - The button is red.
    ///
    /// Note that these values are case-sensitive.
    pub hint_button_style: Option<String>,
    /// The tooltip for a button, if the button has an empty content string.
    pub hint_tool_tip: Option<String>,
}

#[allow(dead_code)]
pub enum ActionsElem {
    Action_(Action),
    Input_(Input),
}

pub struct BindingElem {
    pub icon: Option<Image>,
    pub hero: Option<Image>,
    /// The elements are (text, hint-maxLines).
    pub title: (String, Option<u32>),
    pub elems: Vec<BindingInnerElement>,
}

pub struct Toast {
    pub app_id: String,
    pub binding_elems: Vec<BindingElem>,
    pub duration: Duration,
    pub scenario: Scenario,
    /// The elements are (loop, silent and src).
    pub audio: Option<(bool, bool, SoundSrc)>,
    pub actions: Vec<ActionsElem>,
    pub use_btn_style: bool,
    pub display_timestamp: Option<String>,
}

impl Toast {
    pub const POWERSHELL_APP_ID: &'static str = "{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\
                                                 \\WindowsPowerShell\\v1.0\\powershell.exe";

    pub fn create_notification(&self) -> Result<ToastNotification> {
        let xml: XmlDocument = XmlDocument::new()?;
        let toast_elem: XmlElement = xml.CreateElement(w!("toast"))?;
        if self.use_btn_style {
            toast_elem.SetAttribute(w!("useButtonStyle"), &HSTRING::from(self.use_btn_style.to_string()))?;
        }
        if let Some(timestamp) = &self.display_timestamp {
            toast_elem.SetAttribute(w!("displayTimestamp"), &HSTRING::from(timestamp))?;
        }
        {
            match self.duration {
                Duration::Default => {}
                Duration::Short => {
                    toast_elem.SetAttribute(w!("duration"), w!("short"))?;
                }
                Duration::Long => {
                    toast_elem.SetAttribute(w!("duration"), w!("long"))?;
                }
            };
            match self.scenario {
                Scenario::Default => {}
                Scenario::Alarm => {
                    toast_elem.SetAttribute(w!("scenario"), w!("alarm"))?;
                }
                Scenario::Reminder => {
                    toast_elem.SetAttribute(w!("scenario"), w!("reminder"))?;
                }
                Scenario::IncomingCall => {
                    toast_elem.SetAttribute(w!("scenario"), w!("incomingCall"))?;
                }
                Scenario::Urgent => {
                    toast_elem.SetAttribute(w!("scenario"), w!("urgent"))?;
                }
            };
            let visual_elem: XmlElement = xml.CreateElement(w!("visual"))?;
            {
                for elem in &self.binding_elems {
                    let binding_elem: XmlElement = xml.CreateElement(w!("binding"))?;
                    {
                        binding_elem.SetAttribute(w!("template"), w!("ToastGeneric"))?;
                        if let Some(image) = &elem.icon {
                            add_image_elem_with_placement(image, &xml, &binding_elem, "appLogoOverride", false)?;
                        }
                        if let Some(image) = &elem.hero {
                            add_image_elem_with_placement(image, &xml, &binding_elem, "hero", false)?;
                        }
                        let title_elem: XmlElement = xml.CreateElement(w!("text"))?;
                        {
                            if let Some(value) = &elem.title.1 {
                                title_elem.SetAttribute(w!("hint-maxLines"), &HSTRING::from(value.to_string()))?;
                            }
                            title_elem.SetInnerText(&HSTRING::from(&elem.title.0))?;
                            binding_elem.AppendChild(InParam::owned(IXmlNode::try_from(title_elem)?))?;
                        }
                        for elem in &elem.elems {
                            match elem {
                                BindingInnerElement::Text(text) => {
                                    let text_elem: XmlElement = xml.CreateElement(w!("text"))?;
                                    {
                                        text_elem.SetInnerText(&HSTRING::from(text))?;
                                    }
                                    binding_elem.AppendChild(InParam::owned(IXmlNode::try_from(text_elem)?))?;
                                }
                                BindingInnerElement::Group(ref group_elems) => {
                                    add_group(group_elems, &xml, &binding_elem)?;
                                }
                                BindingInnerElement::Image(ref image) => {
                                    add_image_elem(image, &xml, &binding_elem)?;
                                }
                            }
                        }
                        visual_elem.AppendChild(InParam::owned(IXmlNode::try_from(binding_elem)?))?;
                    }
                }
            }
            toast_elem.AppendChild(InParam::owned(IXmlNode::try_from(visual_elem)?))?;
            add_audio(&self.audio, &xml, &toast_elem)?;
            add_actions(&self.actions, &xml, &toast_elem)?;
        }

        xml.AppendChild(InParam::owned(IXmlNode::try_from(toast_elem)?))?;
        println!("{}", xml.GetXml()?);
        ToastNotification::CreateToastNotification(&xml)
    }

    pub fn show(&self) -> Result<()> {
        let notifier: ToastNotifier = ToastNotificationManager::CreateToastNotifierWithId(&HSTRING::from(&self.app_id))?;
        notifier.Show(&self.create_notification()?)?;
        Ok(())
    }

    pub fn show_with_xml(app_id: &str, xml_content: &str) -> Result<()> {
        let xml: XmlDocument = XmlDocument::new()?;
        xml.LoadXml(&HSTRING::from(xml_content))?;
        let notification: ToastNotification = ToastNotification::CreateToastNotification(&xml)?;
        let notifier: ToastNotifier = ToastNotificationManager::CreateToastNotifierWithId(&HSTRING::from(app_id))?;
        notifier.Show(&notification)?;
        Ok(())
    }
}

fn add_actions(actions_elems: &Vec<ActionsElem>, xml: &XmlDocument, toast_elem: &XmlElement) -> Result<()> {
    if actions_elems.is_empty() {
        return Ok(());
    }
    let actions_elem: XmlElement = xml.CreateElement(w!("actions"))?;
    let (mut action_count, mut input_count) = (0, 0);
    for elem in actions_elems {
        if action_count >= 5 && input_count >= 5 {
            break;
        }
        match elem {
            ActionsElem::Action_(action) => {
                if action_count < 5 {
                    let action_elem: XmlElement = xml.CreateElement(w!("action"))?;
                    action_elem.SetAttribute(w!("content"), &HSTRING::from(&action.content))?;
                    action_elem.SetAttribute(w!("arguments"), &HSTRING::from(&action.arguments))?;
                    match action.activation_type {
                        ActivationType::Foreground => {
                            action_elem.SetAttribute(w!("activationType"), &HSTRING::from("foreground"))?;
                        }
                        ActivationType::Background => {
                            action_elem.SetAttribute(w!("activationType"), &HSTRING::from("background"))?;
                        }
                        ActivationType::Protocol => {
                            action_elem.SetAttribute(w!("activationType"), &HSTRING::from("protocol"))?;
                        }
                    }
                    if matches!(action.placement, ActionPlacement::ContextMenu) {
                        action_elem.SetAttribute(w!("activationType"), w!("contextMenu"))?;
                    }
                    if let Some(value) = &action.image_uri {
                        action_elem.SetAttribute(w!("imageUri"), &HSTRING::from(value))?;
                    }
                    if let Some(value) = &action.hint_input_id {
                        action_elem.SetAttribute(w!("hint-inputId"), &HSTRING::from(value))?;
                    }
                    if let Some(value) = &action.hint_button_style {
                        match value.as_str() {
                            "Success" | "Critical" => {
                                action_elem.SetAttribute(w!("hint-buttonStyle"), &HSTRING::from(value))?;
                            }
                            _ => {}
                        }
                    }
                    if let Some(value) = &action.hint_tool_tip {
                        action_elem.SetAttribute(w!("hint-toolTip"), &HSTRING::from(value))?;
                    }
                    actions_elem.AppendChild(InParam::owned(IXmlNode::try_from(action_elem)?))?;
                    action_count += 1;
                }
            }
            ActionsElem::Input_(input) => {
                if input_count < 5 {
                    let input_elem: XmlElement = xml.CreateElement(w!("input"))?;
                    input_elem.SetAttribute(w!("id"), &HSTRING::from(&input.id))?;
                    let type_ = match &input.type_ {
                        InputType::Text(default) => {
                            if let Some(content) = &input.place_holder_content {
                                input_elem.SetAttribute(w!("placeHolderContent"), &HSTRING::from(content))?;
                            }
                            if let Some(value) = default {
                                input_elem.SetAttribute(w!("defaultInput"), &HSTRING::from(value))?;
                            }
                            Ok::<&str, Error>("text")
                        }
                        InputType::Selection(default_id, selections) => {
                            if let Some(value) = default_id {
                                input_elem.SetAttribute(w!("defaultSelectionBoxItemId"), &HSTRING::from(value))?;
                            }
                            let mut selection_count = 0;
                            for (id, title) in selections {
                                if selection_count >= 5 {
                                    break;
                                }
                                let selection_elem: XmlElement = xml.CreateElement(w!("selection"))?;
                                selection_elem.SetAttribute(w!("id"), &HSTRING::from(id))?;
                                selection_elem.SetAttribute(w!("title"), &HSTRING::from(title))?;
                                input_elem.AppendChild(InParam::owned(IXmlNode::try_from(selection_elem)?))?;
                                selection_count += 1;
                            }
                            Ok::<&str, Error>("selection")
                        }
                    }?;
                    input_elem.SetAttribute(w!("type"), &HSTRING::from(type_))?;
                    if let Some(title) = &input.title {
                        input_elem.SetAttribute(w!("title"), &HSTRING::from(title))?;
                    }
                    actions_elem.AppendChild(InParam::owned(IXmlNode::try_from(input_elem)?))?;
                    input_count += 1;
                }
            }
        }
    }
    toast_elem.AppendChild(InParam::owned(IXmlNode::try_from(actions_elem)?))?;
    Ok(())
}

fn add_image_elem(image: &Image, xml: &XmlDocument, binding_elem: &XmlElement) -> Result<()> {
    add_image_elem_with_placement(image, xml, binding_elem, "", true)
}

fn add_image_elem_with_placement(image: &Image, xml: &XmlDocument, binding_elem: &XmlElement, placement: &str, in_group: bool) -> Result<()> {
    if image.src.is_empty() {
        return Err(Error::new(HRESULT(-1), HSTRING::from("src is empty")));
    }
    let image_elem: XmlElement = xml.CreateElement(w!("image"))?;
    {
        image_elem.SetAttribute(w!("src"), &HSTRING::from(&image.src))?;
        if !placement.is_empty() {
            image_elem.SetAttribute(w!("placement"), &HSTRING::from(placement))?;
        }
        if !image.alt.is_empty() {
            image_elem.SetAttribute(w!("alt"), &HSTRING::from(&image.alt))?;
        }
        if matches!(image.hint_crop, Circle) {
            image_elem.SetAttribute(w!("hint-crop"), w!("circle"))?;
        }
        if in_group {
            match image.hint_align {
                ImageAlign::Default => {}
                ImageAlign::Stretch => {
                    image_elem.SetAttribute(w!("hint-align"), w!("stretch"))?;
                }
                ImageAlign::Left => {
                    image_elem.SetAttribute(w!("hint-align"), w!("left"))?;
                }
                ImageAlign::Center => {
                    image_elem.SetAttribute(w!("hint-align"), w!("center"))?;
                }
                ImageAlign::Right => {
                    image_elem.SetAttribute(w!("hint-align"), w!("right"))?;
                }
            }
        }
        binding_elem.AppendChild(InParam::owned(IXmlNode::try_from(image_elem)?))?;
    }
    Ok(())
}

fn add_in_group_text_elem(text: &InnerText, xml: &XmlDocument, subgroup_elem: &XmlElement) -> Result<()> {
    let text_elem: XmlElement = xml.CreateElement(w!("text"))?;
    {
        if let Some(value) = text.hint_max_lines {
            text_elem.SetAttribute(w!("hint-maxLines"), &HSTRING::from(value.to_string()))?;
        }
        if let Some(value) = text.hint_min_lines {
            text_elem.SetAttribute(w!("hint-minLines"), &HSTRING::from(value.to_string()))?;
        }
        if let Some(value) = text.hint_warp {
            if value {
                text_elem.SetAttribute(w!("hint-warp"), w!("true"))?;
            }
        }
        match text.hint_style {
            TextStyle::Default => {}
            TextStyle::Caption => {
                text_elem.SetAttribute(w!("hint-style"), w!("caption"))?;
            }
            TextStyle::CaptionSubtle => {
                text_elem.SetAttribute(w!("hint-style"), w!("captionSubtle"))?;
            }
            TextStyle::Body => {
                text_elem.SetAttribute(w!("hint-style"), w!("body"))?;
            }
            TextStyle::BodySubtle => {
                text_elem.SetAttribute(w!("hint-style"), w!("bodySubtle"))?;
            }
            TextStyle::Base => {
                text_elem.SetAttribute(w!("hint-style"), w!("base"))?;
            }
            TextStyle::BaseSubtle => {
                text_elem.SetAttribute(w!("hint-style"), w!("baseSubtle"))?;
            }
            TextStyle::Subtitle => {
                text_elem.SetAttribute(w!("hint-style"), w!("subtitle"))?;
            }
            TextStyle::SubtitleSubtle => {
                text_elem.SetAttribute(w!("hint-style"), w!("subtitleSubtle"))?;
            }
            TextStyle::Title => {
                text_elem.SetAttribute(w!("hint-style"), w!("title"))?;
            }
            TextStyle::TitleSubtle => {
                text_elem.SetAttribute(w!("hint-style"), w!("titleSubtle"))?;
            }
            TextStyle::TitleNumeral => {
                text_elem.SetAttribute(w!("hint-style"), w!("titleNumeral"))?;
            }
            TextStyle::Subheader => {
                text_elem.SetAttribute(w!("hint-style"), w!("subheader"))?;
            }
            TextStyle::SubheaderSubtle => {
                text_elem.SetAttribute(w!("hint-style"), w!("subheaderSubtle"))?;
            }
            TextStyle::SubheaderNumeral => {
                text_elem.SetAttribute(w!("hint-style"), w!("subheaderNumeral"))?;
            }
            TextStyle::Header => {
                text_elem.SetAttribute(w!("hint-style"), w!("header"))?;
            }
            TextStyle::HeaderSubtle => {
                text_elem.SetAttribute(w!("hint-style"), w!("headerSubtle"))?;
            }
            TextStyle::HeaderNumeral => {
                text_elem.SetAttribute(w!("hint-style"), w!("headerNumeral"))?;
            }
        }
        match text.hint_align {
            TextAlign::Default => {}
            TextAlign::Auto => {
                text_elem.SetAttribute(w!("hint-align"), w!("auto"))?;
            }
            TextAlign::Left => {
                text_elem.SetAttribute(w!("hint-align"), w!("left"))?;
            }
            TextAlign::Center => {
                text_elem.SetAttribute(w!("hint-align"), w!("center"))?;
            }
            TextAlign::Right => {
                text_elem.SetAttribute(w!("hint-align"), w!("right"))?;
            }
        }
        text_elem.SetInnerText(&HSTRING::from(&text.text))?;
    }
    subgroup_elem.AppendChild(InParam::owned(IXmlNode::try_from(text_elem)?))?;
    Ok(())
}

fn add_group(group_elems: &Vec<Vec<SubgroupElement>>, xml: &XmlDocument, binding_elem: &XmlElement) -> Result<()> {
    if group_elems.is_empty() {
        return Ok(());
    }
    let group_elem: XmlElement = xml.CreateElement(w!("group"))?;
    for subgroup_elems in group_elems.iter()
        .filter(|elem| !elem.is_empty()) {
        let subgroup_elem: XmlElement = xml.CreateElement(w!("subgroup"))?;
        {
            for elem in subgroup_elems {
                match elem {
                    SubgroupElement::Text_(text) => {
                        add_in_group_text_elem(text, &xml, &subgroup_elem)?;
                    }
                    SubgroupElement::Image_(image) => {
                        add_image_elem(image, &xml, &subgroup_elem)?;
                    }
                }
            }
            group_elem.AppendChild(InParam::owned(IXmlNode::try_from(subgroup_elem)?))?;
        }
    }
    binding_elem.AppendChild(InParam::owned(IXmlNode::try_from(group_elem)?))?;
    Ok(())
}

fn add_audio(audio: &Option<(bool, bool, SoundSrc)>, xml: &XmlDocument, toast_elem: &XmlElement) -> Result<()> {
    if let Some((loop_, silent, src)) = audio {
        let audio_elem: XmlElement = xml.CreateElement(w!("audio"))?;
        audio_elem.SetAttribute(w!("loop"), w!("false"))?;
        audio_elem.SetAttribute(w!("silent"), &HSTRING::from(silent.to_string()))?;
        match src {
            SoundSrc::Default => {
                audio_elem.SetAttribute(w!("src"), w!("ms-winsoundevent:Notification.Default"))?;
            }
            SoundSrc::IM => {
                audio_elem.SetAttribute(w!("src"), w!("ms-winsoundevent:Notification.IM"))?;
            }
            SoundSrc::Mail => {
                audio_elem.SetAttribute(w!("src"), w!("ms-winsoundevent:Notification.Mail"))?;
            }
            SoundSrc::Reminder => {
                audio_elem.SetAttribute(w!("src"), w!("ms-winsoundevent:Notification.Reminder"))?;
            }
            SoundSrc::SMS => {
                audio_elem.SetAttribute(w!("src"), w!("ms-winsoundevent:Notification.SMS"))?;
            }
            SoundSrc::Alarm => {
                audio_elem.SetAttribute(w!("src"), w!("ms-winsoundevent:Notification.Looping.Alarm"))?;
                audio_elem.SetAttribute(w!("loop"), &HSTRING::from(loop_.to_string()))?;
            }
            SoundSrc::Alarm2 => {
                audio_elem.SetAttribute(w!("src"), w!("ms-winsoundevent:Notification.Looping.Alarm2"))?;
                audio_elem.SetAttribute(w!("loop"), &HSTRING::from(loop_.to_string()))?;
            }
            SoundSrc::Alarm3 => {
                audio_elem.SetAttribute(w!("src"), w!("ms-winsoundevent:Notification.Looping.Alarm3"))?;
                audio_elem.SetAttribute(w!("loop"), &HSTRING::from(loop_.to_string()))?;
            }
            SoundSrc::Alarm4 => {
                audio_elem.SetAttribute(w!("src"), w!("ms-winsoundevent:Notification.Looping.Alarm4"))?;
                audio_elem.SetAttribute(w!("loop"), &HSTRING::from(loop_.to_string()))?;
            }
            SoundSrc::Alarm5 => {
                audio_elem.SetAttribute(w!("src"), w!("ms-winsoundevent:Notification.Looping.Alarm5"))?;
                audio_elem.SetAttribute(w!("loop"), &HSTRING::from(loop_.to_string()))?;
            }
            SoundSrc::Alarm6 => {
                audio_elem.SetAttribute(w!("src"), w!("ms-winsoundevent:Notification.Looping.Alarm6"))?;
                audio_elem.SetAttribute(w!("loop"), &HSTRING::from(loop_.to_string()))?;
            }
            SoundSrc::Alarm7 => {
                audio_elem.SetAttribute(w!("src"), w!("ms-winsoundevent:Notification.Looping.Alarm7"))?;
                audio_elem.SetAttribute(w!("loop"), &HSTRING::from(loop_.to_string()))?;
            }
            SoundSrc::Alarm8 => {
                audio_elem.SetAttribute(w!("src"), w!("ms-winsoundevent:Notification.Looping.Alarm8"))?;
                audio_elem.SetAttribute(w!("loop"), &HSTRING::from(loop_.to_string()))?;
            }
            SoundSrc::Alarm9 => {
                audio_elem.SetAttribute(w!("src"), w!("ms-winsoundevent:Notification.Looping.Alarm9"))?;
                audio_elem.SetAttribute(w!("loop"), &HSTRING::from(loop_.to_string()))?;
            }
            SoundSrc::Alarm10 => {
                audio_elem.SetAttribute(w!("src"), w!("ms-winsoundevent:Notification.Looping.Alarm10"))?;
                audio_elem.SetAttribute(w!("loop"), &HSTRING::from(loop_.to_string()))?;
            }
            SoundSrc::Call => {
                audio_elem.SetAttribute(w!("src"), w!("ms-winsoundevent:Notification.Looping.Call"))?;
                audio_elem.SetAttribute(w!("loop"), &HSTRING::from(loop_.to_string()))?;
            }
            SoundSrc::Call2 => {
                audio_elem.SetAttribute(w!("src"), w!("ms-winsoundevent:Notification.Looping.Call2"))?;
                audio_elem.SetAttribute(w!("loop"), &HSTRING::from(loop_.to_string()))?;
            }
            SoundSrc::Call3 => {
                audio_elem.SetAttribute(w!("src"), w!("ms-winsoundevent:Notification.Looping.Call3"))?;
                audio_elem.SetAttribute(w!("loop"), &HSTRING::from(loop_.to_string()))?;
            }
            SoundSrc::Call4 => {
                audio_elem.SetAttribute(w!("src"), w!("ms-winsoundevent:Notification.Looping.Call4"))?;
                audio_elem.SetAttribute(w!("loop"), &HSTRING::from(loop_.to_string()))?;
            }
            SoundSrc::Call5 => {
                audio_elem.SetAttribute(w!("src"), w!("ms-winsoundevent:Notification.Looping.Call5"))?;
                audio_elem.SetAttribute(w!("loop"), &HSTRING::from(loop_.to_string()))?;
            }
            SoundSrc::Call6 => {
                audio_elem.SetAttribute(w!("src"), w!("ms-winsoundevent:Notification.Looping.Call6"))?;
                audio_elem.SetAttribute(w!("loop"), &HSTRING::from(loop_.to_string()))?;
            }
            SoundSrc::Call7 => {
                audio_elem.SetAttribute(w!("src"), w!("ms-winsoundevent:Notification.Looping.Call7"))?;
                audio_elem.SetAttribute(w!("loop"), &HSTRING::from(loop_.to_string()))?;
            }
            SoundSrc::Call8 => {
                audio_elem.SetAttribute(w!("src"), w!("ms-winsoundevent:Notification.Looping.Call8"))?;
                audio_elem.SetAttribute(w!("loop"), &HSTRING::from(loop_.to_string()))?;
            }
            SoundSrc::Call9 => {
                audio_elem.SetAttribute(w!("src"), w!("ms-winsoundevent:Notification.Looping.Call9"))?;
                audio_elem.SetAttribute(w!("loop"), &HSTRING::from(loop_.to_string()))?;
            }
            SoundSrc::Call10 => {
                audio_elem.SetAttribute(w!("src"), w!("ms-winsoundevent:Notification.Looping.Call10"))?;
                audio_elem.SetAttribute(w!("loop"), &HSTRING::from(loop_.to_string()))?;
            }
        }
        toast_elem.AppendChild(InParam::owned(IXmlNode::try_from(audio_elem)?))?;
    }
    Ok(())
}
