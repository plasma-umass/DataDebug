// -*- C++ -*-

#ifndef FYSHUFFLE_H_
#define FYSHUFFLE_H_

#include <stdlib.h>
#include <vector>
using namespace std;

namespace fyshuffle {

  template <class TYPE>
  void inplace (vector<TYPE>& vec)
  {
    for (auto i = vec.size()-1; i > 0; i--) {
      auto j = lrand48() % (i+1);
      swap (vec[i], vec[j]);
    }
  }

  template <class TYPE>
  void transform (const vector<TYPE>& in,
		  vector<TYPE>& out)
  {
    out[0] = in[0];
    for (auto i = 1; i < in.size(); i++) {
      auto j = lrand48() % (i+1);
      out[i] = out[j];
      out[j] = in[i];
    }
  }

}


#endif
