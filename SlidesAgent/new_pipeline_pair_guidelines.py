from __future__ import annotations

from SlidesAgent.new_pipeline import build_arg_parser, run_pipeline


def main() -> None:
    parser = build_arg_parser(
        description='Experimental SlideGen pipeline using pair-derived paper-slide guidelines.',
        include_pair_guideline_args=True,
    )
    args = parser.parse_args()

    if getattr(args, "use_author_preferences", False):
        print("[pair-guidelines] Ignoring --use_author_preferences for this experimental entrypoint.")
        args.use_author_preferences = False

    args.use_pair_guidelines = True
    args.output_variant_suffix = "_pair_guidelines"
    args.output_folder_suffix = ""
    run_pipeline(args)


if __name__ == "__main__":
    main()
