
## Pull Requests and Contributing to rust_rust_xlsxwriter

All patches and pull requests are welcome but must start with an issue tracker.


### Getting Started

If the change is small such as a documentation or syntax fix then submit the change without the steps below.

For anything else follow the steps below:

1. Pull requests and new feature proposals must start with an [Issue Tracker](https://github.com/jmcnamara/rust_xlsxwriter/issues). This serves as the focal point for the design discussion. Start with the Question or Feature Request template.
2. Describe what you plan to do. If there are API changes or additions add some pseudo-code to demonstrate them.
3. Wait for a response (it should come quickly) before submitting a PR.

Once there has been some discussion in the Issue Tracker you can prepare the PR as follows:

1. Fork the repository.
2. Run all the tests to make sure the current code work on your system using `cargo test`.
3. Create a feature branch for your new feature.
4. Code style is `rustfmt`.
5. Rebase changes into one commit unless it requires separate logical steps, see below.


### Writing Tests

Where possible add unit tests in the `src/*.rs` files.

Integration tests in the `tests` folder are harder to create since the generally require an input Excel 2007
file to test against. The maintainer can help with this and it can be discussed in the Issue Tracker.

### Example programs

If applicable add an example program to the `examples` directory.

### Copyright and License

Copyright remains with the original author. Do not include additional copyright claims or Licensing requirements. GitHub and the `git` repository will record your contribution.


### Submitting the Pull Request

Follow the commit message style of the commit log which is roughly Linux kernel style.

If your change involves several incremental `git` commits then `rebase` or `squash` them onto another branch so that the Pull Request is a single commit or a small number of logical commits.

Push your changes to GitHub and submit the Pull Request with a hash link to the to the Issue tracker that was opened above.
