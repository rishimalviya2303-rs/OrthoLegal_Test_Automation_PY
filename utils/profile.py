from playwright.sync_api import Locator, Page, expect


def dismiss_profile_picture_popup_if_present(page: Page) -> None:
    """Cancel the optional photo dialog, including one arriving during navigation."""
    popup = page.get_by_role("dialog").filter(
        has=page.get_by_role("heading", name="Add a photo", exact=True)
    )

    def cancel_photo_dialog(dialog: Locator) -> None:
        cancel = dialog.get_by_role("button", name="Cancel", exact=True)
        expect(cancel, "The profile-picture popup must offer Cancel.").to_be_visible()
        cancel.click()
        expect(dialog, "The profile-picture popup must close after Cancel.").not_to_be_visible()

    # Actionability checks catch a dialog arriving after the immediate check.
    # An absent popup adds no timeout and never prevents normal navigation.
    page.add_locator_handler(popup, cancel_photo_dialog)
    if popup.is_visible():
        cancel_photo_dialog(popup)
