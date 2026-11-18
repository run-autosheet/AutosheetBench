// The reasoning beside a model's name, unless it repeats it: a product that
// picks its own model picks its reasoning too, which reads "Auto (Auto)".
// Its own module, so a card can name a model without the benchmark's data.
export const shownReasoning = ({ modelName, reasoning }) =>
  reasoning === modelName ? null : reasoning
