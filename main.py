async def fill_arbk_nui(page, nui: str):
    """
    Gjen fushën zyrtare:
    'Numri unik identifikues ose numri i biznesit'
    dhe e plotëson me numrin e kërkuar.
    """

    # 1. Provo sipas placeholder-it real të ARBK
    selectors = [
        'input[placeholder*="Numri unik identifikues"]',
        'input[placeholder*="numri i biznesit"]',
        'input[placeholder*="Numri unik"]'
    ]

    for selector in selectors:
        locator = page.locator(selector)

        if await locator.count() > 0:
            await locator.first.fill(nui)
            return True

    # 2. Provo sipas tekstit/label-it pranë inputit
    labels = page.locator("label")

    for i in range(await labels.count()):
        label = labels.nth(i)

        try:
            text = (await label.inner_text()).lower()

            if (
                "numri unik identifikues" in text
                or "numri i biznesit" in text
            ):
                target = await label.get_attribute("for")

                if target:
                    field = page.locator(f"#{target}")

                    if await field.count() > 0:
                        await field.fill(nui)
                        return True

        except Exception:
            continue

    # 3. Fallback: kontrollo input-et e dukshme
    inputs = page.locator("input")

    for i in range(await inputs.count()):
        field = inputs.nth(i)

        try:
            if not await field.is_visible():
                continue

            placeholder = (
                await field.get_attribute("placeholder") or ""
            ).lower()

            aria = (
                await field.get_attribute("aria-label") or ""
            ).lower()

            combined = f"{placeholder} {aria}"

            if (
                "numri unik identifikues" in combined
                or "numri i biznesit" in combined
            ):
                await field.fill(nui)
                return True

        except Exception:
            continue

    return False
