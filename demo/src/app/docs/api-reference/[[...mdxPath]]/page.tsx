import { generateStaticParamsFor, importPage } from "nextra/pages";
import { useMDXComponents } from "../../../../../mdx-components";

export const generateStaticParams = generateStaticParamsFor("mdxPath");

export async function generateMetadata(props: PageProps<"/docs/api-reference/[[...mdxPath]]">) {
  const params = await props.params;
  const { metadata } = await importPage(params.mdxPath);
  return metadata;
}

const Wrapper = useMDXComponents().wrapper;

export default async function ApiReferencePage(
  props: PageProps<"/docs/api-reference/[[...mdxPath]]">,
) {
  if (Wrapper === undefined) {
    throw new Error("Nextra documentation wrapper is unavailable");
  }

  const params = await props.params;
  const { default: MDXContent, toc, metadata, sourceCode } = await importPage(params.mdxPath);

  return (
    <Wrapper toc={toc} metadata={metadata} sourceCode={sourceCode}>
      <MDXContent {...props} params={params} />
    </Wrapper>
  );
}
