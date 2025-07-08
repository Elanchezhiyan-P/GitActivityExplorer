using EnvDTE;
using EnvDTE80;
using GitActivityExplorer;
using LibGit2Sharp;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.Win32;
using OxyPlot;
using OxyPlot.Axes;
using OxyPlot.Series;
using OxyPlot.Wpf;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Threading;
using DataPoint = OxyPlot.DataPoint;

namespace GitActivityExplorer
{
    public partial class GitActivityWindowControl : UserControl, IDisposable
    {
        private Repository repo;
        private IList<Commit> commits;
        private string repoPath;
        private uint solutionEventsCookie;
        private IVsSolution solution;
        private SolutionEventsHandler eventsHandler;
        private int selectedCommitIndex = -1;

        public class CommitViewModel
        {
            public string Author { get; set; }
            public string Message { get; set; }
            public string Date { get; set; }
            public string CommitId { get; set; }
            public Commit OriginalCommit { get; set; }
        }

        public class FileChangeViewModel
        {
            public string Status { get; set; }
            public string Path { get; set; }
            public Brush RowColor
            {
                get
                {
                    switch (Status)
                    {
                        case "Added": return Brushes.LightGreen;
                        case "Deleted": return Brushes.IndianRed;
                        case "Modified": return Brushes.LightBlue;
                        case "Renamed": return Brushes.LightGoldenrodYellow;
                        default: return Brushes.Transparent;
                    }
                }
            }
        }

        public GitActivityWindowControl()
        {
            InitializeComponent();
            SubscribeToSolutionEvents();
            LoadRepository();
        }

        private void SubscribeToSolutionEvents()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            eventsHandler = new SolutionEventsHandler(this);
            solution = (IVsSolution)Package.GetGlobalService(typeof(SVsSolution));
            solution?.AdviseSolutionEvents(eventsHandler, out solutionEventsCookie);
        }

        private void ShowLoading(bool isLoading)
        {
            LoadingSpinner.Visibility = isLoading ? Visibility.Visible : Visibility.Collapsed;
        }

        private void RunGitCommand(string arguments, string title)
        {
            try
            {
                var psi = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = "git",
                    Arguments = arguments,
                    WorkingDirectory = repoPath,
                    RedirectStandardOutput = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };
                var process = System.Diagnostics.Process.Start(psi);
                string output = process.StandardOutput.ReadToEnd();
                process.WaitForExit();
                MessageBox.Show(output, title);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}");
            }
        }

        private string GetSolutionDirectory()
        {
            try
            {
                var dte = (DTE2)ServiceProvider.GlobalProvider.GetService(typeof(DTE));
                var solutionPath = dte?.Solution?.FullName;
                return !string.IsNullOrEmpty(solutionPath) ? Path.GetDirectoryName(solutionPath) : null;
            }
            catch { return null; }
        }

        public void LoadRepository()
        {
            FileChangesTable.ItemsSource = null;
            StatusTextBlock.Text = string.Empty;
            CommitTable.ItemsSource = null;
            ShowLoading(true);

            _ = Dispatcher.InvokeAsync(() =>
            {
                try
                {
                    var solutionDir = GetSolutionDirectory();
                    if (string.IsNullOrEmpty(solutionDir) || !Directory.Exists(solutionDir))
                    {
                        StatusTextBlock.Text = "❌ No solution is currently loaded.";
                        return;
                    }

                    repoPath = Repository.Discover(solutionDir);
                    if (string.IsNullOrEmpty(repoPath))
                    {
                        StatusTextBlock.Text = "❌ No Git repository found in solution path.";
                        return;
                    }

                    repo = new Repository(repoPath);
                    BranchSelector.ItemsSource = repo.Branches.Where(b => !b.IsRemote).Select(b => b.FriendlyName).ToList();
                    BranchSelector.SelectedItem = repo.Head.FriendlyName;
                    LoadGitCommits(repo.Head.FriendlyName);
                    StatusTextBlock.Text = $"✅ Loaded repository: {repoPath}";
                }
                catch (Exception ex)
                {
                    StatusTextBlock.Text = $"❌ Error opening repository: {ex.Message}";
                }
                finally { ShowLoading(false); }
            }, DispatcherPriority.Background);
        }

        private void LoadGitCommits(string branchName)
        {
            if (repo == null) return;
            var branch = repo.Branches[branchName];
            commits = branch.Commits.Take(1000).ToList();

            CommitTable.ItemsSource = commits.Select(c => new CommitViewModel
            {
                Author = c.Author.Name,
                Message = c.MessageShort,
                Date = c.Author.When.LocalDateTime.ToString("yyyy-MM-dd HH:mm"),
                CommitId = c.Sha.Substring(0, 8),
                OriginalCommit = c
            }).ToList();

            selectedCommitIndex = -1;
            ShowInsights(commits);
        }

        private void ShowInsights(IEnumerable<Commit> commitList)
        {
            var commitsPerDay = commitList.GroupBy(c => c.Author.When.Date).Select(g => new { Date = g.Key, Count = g.Count() }).ToList();
            var commitsPerAuthor = commitList.GroupBy(c => c.Author.Name).Select(g => new { Author = g.Key, Count = g.Count() }).ToList();
            var commitsPerFile = new Dictionary<string, int>();

            foreach (var commit in commitList)
            {
                var parent = commit.Parents.FirstOrDefault();
                if (parent == null) continue;
                var patch = repo.Diff.Compare<Patch>(parent.Tree, commit.Tree);
                foreach (var change in patch)
                {
                    if (!commitsPerFile.ContainsKey(change.Path))
                        commitsPerFile[change.Path] = 0;
                    commitsPerFile[change.Path]++;
                }
            }

            var mostActiveAuthor = commitsPerAuthor.OrderByDescending(a => a.Count).FirstOrDefault()?.Author;
            var busiestDay = commitsPerDay.OrderByDescending(c => c.Count).FirstOrDefault();
            var topModifiedFile = commitsPerFile.OrderByDescending(f => f.Value).FirstOrDefault().Key;
            var longestMessage = commitList.OrderByDescending(c => c.Message.Length).FirstOrDefault()?.Message;

            StatusTextBlock.Text = string.Empty;
            StatusTextBlock.Text += $"\n👨‍💻 Most Active Author: {mostActiveAuthor}";
            StatusTextBlock.Text += $"\n📅 Busiest Day: {busiestDay?.Date:yyyy-MM-dd} ({busiestDay?.Count} commits)";
            StatusTextBlock.Text += $"\n📁 Top Modified File: {topModifiedFile}";
            StatusTextBlock.Text += $"\n📝 Longest Message: {longestMessage?.Substring(0, Math.Min(250, longestMessage.Length))}...";

            PlotCommits(commitsPerDay.Select(x => (x.Date, x.Count)).ToList());
        }

        private void button1_Click(object sender, RoutedEventArgs e) => LoadRepository();

        private void BranchSelector_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (BranchSelector.SelectedItem is string branchName)
                LoadGitCommits(branchName);
        }

        private void CommitTable_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            FileChangesTable.ItemsSource = null;

            if (CommitTable.SelectedItem is CommitViewModel selected && repo != null)
            {
                selectedCommitIndex = commits.IndexOf(selected.OriginalCommit);
                var commit = selected.OriginalCommit;
                var parent = commit.Parents.FirstOrDefault();
                if (parent == null) return;

                var patch = repo.Diff.Compare<Patch>(parent.Tree, commit.Tree);
                FileChangesTable.ItemsSource = patch.Select(p => new FileChangeViewModel
                {
                    Status = p.Status.ToString(),
                    Path = p.Path
                }).ToList();
            }
        }

        private void FileChangesTable_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (FileChangesTable.SelectedItem is FileChangeViewModel selectedFile)
            {
                MessageBox.Show($"Diff Preview\nStatus: {selectedFile.Status}\nPath: {selectedFile.Path}", "Diff Preview", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void ExportCSV_Click(object sender, RoutedEventArgs e)
        {
            if (commits == null || !commits.Any()) return;
            var dialog = new SaveFileDialog { Filter = "CSV files (*.csv)|*.csv", DefaultExt = ".csv", FileName = "commits.csv", Title = "Save Commit Data as CSV" };
            if (dialog.ShowDialog() == true)
            {
                var lines = commits.Select(c => $"\"{c.Author.Name}\",\"{c.MessageShort}\",\"{c.Author.When.LocalDateTime:yyyy-MM-dd HH:mm}\",\"{c.Sha.Substring(0, 8)}\"");
                File.WriteAllLines(dialog.FileName, lines);
                MessageBox.Show($"Exported to {dialog.FileName}");
            }
        }

        private void ExportJSON_Click(object sender, RoutedEventArgs e)
        {
            if (commits == null || !commits.Any()) return;
            var dialog = new SaveFileDialog { Filter = "JSON files (*.json)|*.json", DefaultExt = ".json", FileName = "commits.json", Title = "Save Commit Data as JSON" };
            if (dialog.ShowDialog() == true)
            {
                var json = JsonSerializer.Serialize(commits.Select(c => new
                {
                    Author = c.Author.Name,
                    Message = c.MessageShort,
                    Date = c.Author.When.LocalDateTime,
                    CommitId = c.Sha
                }), new JsonSerializerOptions { WriteIndented = true });

                File.WriteAllText(dialog.FileName, json);
                MessageBox.Show($"Exported to {dialog.FileName}");
            }
        }

        private void GitPull_Click(object sender, RoutedEventArgs e) => RunGitCommand("pull", "Git Pull Output");

        private void GitFetch_Click(object sender, RoutedEventArgs e) => RunGitCommand("fetch", "Git Fetch Output");

        private void PrevCommit_Click(object sender, RoutedEventArgs e)
        {
            if (commits == null || commits.Count == 0) return;

            selectedCommitIndex = (selectedCommitIndex - 1 + commits.Count) % commits.Count;
            CommitTable.SelectedIndex = selectedCommitIndex;
            CommitTable.ScrollIntoView(CommitTable.SelectedItem);
        }

        private void NextCommit_Click(object sender, RoutedEventArgs e)
        {
            if (commits == null || commits.Count == 0) return;

            selectedCommitIndex = (selectedCommitIndex + 1) % commits.Count;
            CommitTable.SelectedIndex = selectedCommitIndex;
            CommitTable.ScrollIntoView(CommitTable.SelectedItem);
        }

        private void PlotCommits(List<(DateTime Date, int Count)> commitsPerDay)
        {
            var series = new LineSeries
            {
                Title = "Commits",
                Color = OxyColors.ForestGreen,
                MarkerType = MarkerType.Circle,
                TrackerFormatString = "Date: {2:MM-dd}\nCommits: {4}"
            };

            foreach (var point in commitsPerDay.OrderBy(p => p.Date))
            {
                series.Points.Add(new DataPoint(DateTimeAxis.ToDouble(point.Date), point.Count));
            }

            var model = new PlotModel
            {
                Title = "Commits per Day",
                TextColor = OxyColors.Black,
                PlotAreaBorderColor = OxyColors.Gray,
                DefaultFont = "Segoe UI",
                TitleColor = OxyColors.Black
            };

            model.Series.Add(series);
            model.Axes.Add(new DateTimeAxis
            {
                Position = AxisPosition.Bottom,
                StringFormat = "MM-dd",
                MajorGridlineStyle = LineStyle.Solid,
                Title = "Date"
            });
            model.Axes.Add(new LinearAxis
            {
                Position = AxisPosition.Left,
                Title = "Commits",
                MajorGridlineStyle = LineStyle.Solid,
                MinimumPadding = 0.1,
                MaximumPadding = 0.1
            });

            CommitPlot.Model = model;
        }

        public void Dispose()
        {
            if (solution != null && solutionEventsCookie != 0)
            {
                solution.UnadviseSolutionEvents(solutionEventsCookie);
                solutionEventsCookie = 0;
            }
        }
    }
}