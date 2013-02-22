
template <class TYPE>
void computeOneBootstrap (const vector<TYPE>& mixsource,
			  long M,
			  long N,
			  float& bootstrap)
{
  vector<TYPE> mix;
  mix.resize (M + N);
  
  // Shuffle mix -- this will let us perform sampling without
  // replacement efficiently.
  fyshuffle::transform (mixsource, mix);

  // Now we take the first M as "f", and the next N as "g".
  // Save the absolute difference of their means.
  float sum1 = 0;
  for (auto i = 0; i < M; i++) {
    sum1 += mix[i];
  }
  float sum2 = 0;
  for (auto i = 0; i < N; i++) {
    sum2 += mix[M+i];
  }
  bootstrap = (sum1/(float) M) - (sum2/(float) N);

  //  cout << "# boot avg = " << sum1/M << ", " << sum2 / N << endl << ": diff = " << bootstrap << endl;
}


template <class TYPE>
bool significantDifference (const float significanceLevel,
			    const vector<TYPE>& f,
			    const vector<TYPE>& g,
			    unsigned long NumBootstraps = NBOOTSTRAPS)
{
  assert (significanceLevel > 0.0);
  assert (significanceLevel < 1.0);

  // Compute the original difference in means.
  auto M = f.size();
  auto N = g.size();
  auto originalMeanDiff = (float) average (f) - (float) average (g);

  // Combine both vectors.
  vector<TYPE> combined;
  combined.resize (M + N);
  int index = 0;
  for (auto const& x : f) {
    combined[index++] = x;
  }
  assert (index == M);
  for (auto const& x : g) {
    combined[index++] = x;
  }
  assert (index == M + N);

  // Build up the bootstrap of averages.
  vector<float> bootstrap;
  bootstrap.resize (NumBootstraps);

  for (auto i = 0; i < NumBootstraps; i++) {
    computeOneBootstrap (combined, M, N, bootstrap[i]);
    // cout << "# avg bootstrap" << endl;
    // cout << bootstrap[i] << endl;
  }
      
  // Now check to see whether the original mean is outside the
  // confidence interval.
  sort (bootstrap.begin(), bootstrap.end());

  // Find the left and right intervals.
  int leftInterval = floor(significanceLevel / 2.0 * NumBootstraps);
  int rightInterval = ceil((1.0 - significanceLevel / 2.0) * NumBootstraps);

  cout << "# originalMeanDiff = " << originalMeanDiff << endl;
  cout << "# interval = ["
       << bootstrap[leftInterval] << ","
       << bootstrap[rightInterval] << "]" << endl;

  bool isOutside = ((originalMeanDiff < bootstrap[leftInterval]) ||
  		    (originalMeanDiff > bootstrap[rightInterval]));

  return isOutside;
}


template <class TYPE>
bool significant (const int k,
		  const vector<TYPE>& original,
		  const vector<TYPE>& bootOriginal,
		  bool& result)
{
  vector<TYPE> b;
  b.resize (NELEMENTS);

  // Build a bootstrap distribution WITHOUT index k.
  vector<TYPE> bootWithout;
  bootWithout.resize (NBOOTSTRAPS);
  for (long i = 0; i < NBOOTSTRAPS; i++) {
    bootstrap::exclusive (k, original, b);
    bootWithout[i] = poly(b) / (float) NELEMENTS;
    //    cout << "# boot without" << endl;
    //    cout << bootWithout[i] << endl;
  }
  // Now check to see if there's a significant difference in the
  // distribution means.

  assert (bootOriginal.size() == NBOOTSTRAPS);
  assert (bootWithout.size() == NBOOTSTRAPS);
  result = significantDifference (ALPHA, bootOriginal, bootWithout);


  cout << "# significant difference at " << (1.0-ALPHA) << " level? ";
  if (result) { cout << "YES"; } else { cout << "NO"; }
  cout << endl;
  cout << "# results for index " << k << endl;
  cout << "# avg with = " << average (bootOriginal) << endl;
  cout << "# avg without = " << average (bootWithout) << endl;
  return result;
}
