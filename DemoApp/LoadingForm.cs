using System;
using System.Drawing;
using System.Windows.Forms;

namespace DemoApp
{
    internal sealed class LoadingForm : Form
    {
        private readonly Label messageLabel;
        private readonly ProgressBar progressBar;

        public LoadingForm()
        {
            FormBorderStyle = FormBorderStyle.None;
            StartPosition = FormStartPosition.CenterScreen;
            BackColor = Color.White;
            Size = new Size(360, 120);
            TopMost = true;
            ShowInTaskbar = false;
            DoubleBuffered = true;

            messageLabel = new Label
            {
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Segoe UI", 10F, FontStyle.Regular),
                Text = "Loading Envista Turntable…",
                Padding = new Padding(0, 24, 0, 12)
            };

            progressBar = new ProgressBar
            {
                Dock = DockStyle.Bottom,
                Height = 8,
                Style = ProgressBarStyle.Marquee,
                MarqueeAnimationSpeed = 35
            };

            Controls.Add(messageLabel);
            Controls.Add(progressBar);
        }

        protected override bool ShowWithoutActivation => true;
    }
}
