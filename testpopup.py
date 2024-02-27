from plyer import notification

# ignore pylance issues
notification.notify(
    title='Your Title',
    message='Your Message',
    app_name='Your App Name',
    app_icon='favicon.ico',
    timeout=10,
    ticker='Your Ticker',
    toast=False,
    hints={}
)
